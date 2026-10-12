"""A small Microsoft Graph client for mail.  No ADK code in here on purpose.

Keeping "talking to Outlook" separate from "talking to the LLM" means:
  * we can test this file with a fake HTTP server (no real mailbox needed),
  * the tools file stays tiny and easy to read.

Graph docs used:
  List messages : GET  /me/mailFolders/inbox/messages
  Get message   : GET  /me/messages/{id}
  Send mail     : POST /me/sendMail
"""

import re
import time
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import httpx

from . import auth, config

# Only ask Graph for the fields we actually use (faster, smaller).
LIST_FIELDS = "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments"
FULL_FIELDS = LIST_FIELDS + ",body,toRecipients,ccRecipients"


class GraphError(Exception):
    """Any problem talking to Graph. The message is written for humans/LLM."""


# ---------------------------------------------------------------- time helpers
def local_tz():
    """The user's timezone: USER_TIMEZONE if set, else this computer's zone."""
    import os

    name = os.getenv("USER_TIMEZONE")
    if name:
        return ZoneInfo(name)
    return datetime.now().astimezone().tzinfo


def to_local_string(graph_utc_time: str) -> str:
    """'2026-10-09T10:15:00Z' -> '2026-10-09 15:45 IST' (in the user's zone)."""
    dt = datetime.fromisoformat(graph_utc_time.replace("Z", "+00:00"))
    return dt.astimezone(local_tz()).strftime("%Y-%m-%d %H:%M %Z")


def day_bounds_utc(start_date: str, end_date: str) -> tuple[str, str]:
    """Turn local dates 'YYYY-MM-DD'..'YYYY-MM-DD' (inclusive) into UTC ISO times.

    Example (IST, +05:30): 2026-10-09..2026-10-09 ->
      start = 2026-10-08T18:30:00Z, end = 2026-10-09T18:30:00Z
    """
    tz = local_tz()
    start = datetime.strptime(start_date, "%Y-%m-%d").replace(tzinfo=tz)
    end = datetime.strptime(end_date, "%Y-%m-%d").replace(tzinfo=tz) + timedelta(days=1)
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    return (
        start.astimezone(timezone.utc).strftime(fmt),
        end.astimezone(timezone.utc).strftime(fmt),
    )


# ------------------------------------------------------------- HTTP plumbing
def _request(method: str, path: str, params=None, json=None, extra_headers=None):
    """One place for auth header, error handling, and one retry on throttling."""
    headers = {
        "Authorization": f"Bearer {auth.get_access_token()}",
        # Ask for plain-text bodies instead of messy HTML.
        "Prefer": 'outlook.body-content-type="text"',
    }
    if extra_headers:
        headers.update(extra_headers)

    url = config.GRAPH_BASE_URL + path
    for attempt in (1, 2):
        try:
            resp = httpx.request(
                method, url, params=params, json=json, headers=headers, timeout=30
            )
        except httpx.HTTPError as exc:
            raise GraphError(f"Network problem talking to Microsoft Graph: {exc}")

        if resp.status_code == 429 and attempt == 1:  # throttled: wait, retry once
            time.sleep(min(int(resp.headers.get("Retry-After", "2")), 10))
            continue
        break

    if resp.status_code in (200, 201, 202, 204):
        return resp.json() if resp.content else {}
    if resp.status_code == 401:
        raise GraphError("Microsoft rejected the token (401). Run `python login.py` again.")
    if resp.status_code == 403:
        raise GraphError(
            "Permission denied (403). Check that Mail.Read / Mail.Send are added to "
            "the app registration and that you accepted them at login."
        )
    raise GraphError(f"Graph error {resp.status_code}: {resp.text[:300]}")


# ------------------------------------------------------------ shaping results
def _clean_text(text: str) -> str:
    text = re.sub(r"\r\n?", "\n", text or "")
    text = re.sub(r"\n{3,}", "\n\n", text)  # collapse big blank gaps
    text = text.strip()
    if len(text) > config.BODY_CHAR_LIMIT:
        text = text[: config.BODY_CHAR_LIMIT] + "\n...[truncated]"
    return text


def _address(person: dict | None) -> tuple[str, str]:
    box = (person or {}).get("emailAddress", {})
    return box.get("name", ""), box.get("address", "")


def simplify(msg: dict, full: bool = False) -> dict:
    """Graph's big JSON -> the few fields the LLM needs."""
    name, address = _address(msg.get("from"))
    out = {
        "id": msg.get("id"),
        "from_name": name,
        "from_email": address,
        "subject": msg.get("subject") or "(no subject)",
        "received": to_local_string(msg["receivedDateTime"]),
        "is_read": msg.get("isRead"),
        "has_attachments": msg.get("hasAttachments"),
    }
    if full:
        out["to"] = [_address(r)[1] for r in msg.get("toRecipients", [])]
        out["cc"] = [_address(r)[1] for r in msg.get("ccRecipients", [])]
        out["body"] = _clean_text(msg.get("body", {}).get("content", ""))
    else:
        out["preview"] = _clean_text(msg.get("bodyPreview", ""))[:300]
    return out


def _newest_first(items: list[dict]) -> list[dict]:
    return sorted(items, key=lambda m: m["receivedDateTime"], reverse=True)


# ---------------------------------------------------------------- public API
def get_profile() -> dict:
    data = _request("GET", "/me", params={"$select": "displayName,mail,userPrincipalName"})
    return {
        "name": data.get("displayName"),
        "email": data.get("mail") or data.get("userPrincipalName"),
    }


def list_recent(count: int = 5) -> list[dict]:
    count = max(1, min(count, config.MAX_EMAILS_PER_CALL))
    data = _request(
        "GET",
        "/me/mailFolders/inbox/messages",
        params={"$top": count, "$orderby": "receivedDateTime desc", "$select": LIST_FIELDS},
    )
    return [simplify(m) for m in data.get("value", [])]


def list_between(start_date: str, end_date: str, count: int = 10) -> list[dict]:
    """Emails received from start_date to end_date (both inclusive, local dates)."""
    count = max(1, min(count, config.MAX_EMAILS_PER_CALL))
    start_utc, end_utc = day_bounds_utc(start_date, end_date)
    data = _request(
        "GET",
        "/me/mailFolders/inbox/messages",
        params={
            # NOTE: Graph needs the $orderby property to also appear in $filter.
            "$filter": f"receivedDateTime ge {start_utc} and receivedDateTime lt {end_utc}",
            "$orderby": "receivedDateTime desc",
            "$top": count,
            "$select": LIST_FIELDS,
        },
    )
    return [simplify(m) for m in data.get("value", [])]


def find_from_sender(sender: str, count: int = 10) -> list[dict]:
    """Emails whose sender name or address contains `sender` (case-insensitive).

    Simple approach: take the most recent SCAN_WINDOW emails and filter in
    Python.  (Graph's server-side filter on `from` cannot be combined with
    sorting by date, so this is the simplest reliable option.)
    """
    count = max(1, min(count, config.MAX_EMAILS_PER_CALL))
    data = _request(
        "GET",
        "/me/mailFolders/inbox/messages",
        params={
            "$top": config.SCAN_WINDOW,
            "$orderby": "receivedDateTime desc",
            "$select": LIST_FIELDS,
        },
    )
    needle = sender.strip().lower()
    hits = []
    for m in data.get("value", []):
        name, address = _address(m.get("from"))
        if needle in name.lower() or needle in address.lower():
            hits.append(simplify(m))
    return hits[:count]


def search(query: str, count: int = 10) -> list[dict]:
    """Keyword search over subject/body/sender using Graph's $search (KQL)."""
    count = max(1, min(count, config.MAX_EMAILS_PER_CALL))
    clean = query.replace('"', " ").strip()
    data = _request(
        "GET",
        "/me/messages",
        # $search value must be wrapped in double quotes.  $search cannot be
        # combined with $orderby, so we sort ourselves below.
        params={"$search": f'"{clean}"', "$top": count, "$select": LIST_FIELDS},
    )
    return [simplify(m) for m in _newest_first(data.get("value", []))]


def get_message(message_id: str) -> dict:
    data = _request("GET", f"/me/messages/{message_id}", params={"$select": FULL_FIELDS})
    return simplify(data, full=True)


def send_mail(to: list[str], subject: str, body: str, cc: list[str] | None = None) -> None:
    def recipients(addresses):
        return [{"emailAddress": {"address": a.strip()}} for a in addresses]

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": recipients(to),
            "ccRecipients": recipients(cc or []),
        },
        "saveToSentItems": True,
    }
    _request("POST", "/me/sendMail", json=payload)  # 202 Accepted, empty body
