"""The tools the Gemini agent can call.

How ADK uses these functions
----------------------------
ADK reads each function's NAME, TYPE HINTS and DOCSTRING and shows them to the
model as a menu of tools.  So the docstring is not decoration: it is the
instruction manual the model reads to decide WHEN and HOW to call the tool.

Rules we follow (from ADK docs):
  * simple types only (str, int, list[str]),
  * no default-heavy / exotic parameters,
  * return a dict with a "status" key ("success" or "error"),
  * never raise to the model - return the error text so it can explain it.
"""

import re

from . import graph_client
from .auth import AuthError
from .graph_client import GraphError

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _ok(**data) -> dict:
    return {"status": "success", **data}


def _fail(exc: Exception) -> dict:
    return {"status": "error", "message": str(exc)}


def get_latest_emails(count: int = 5) -> dict:
    """Get the most recent emails from the inbox, newest first.

    Use for: "who sent the latest email?", "show my last 5 emails",
    "summarize the latest email" (then call read_email for the full body).

    Args:
        count: How many emails to return (1-25). Default 5.

    Returns:
        dict with status and a list "emails". Each email has: id, from_name,
        from_email, subject, received (local time), is_read, preview.
    """
    try:
        return _ok(emails=graph_client.list_recent(count))
    except (AuthError, GraphError) as exc:
        return _fail(exc)


def get_emails_by_date(start_date: str, end_date: str = "", count: int = 10) -> dict:
    """Get inbox emails received on a date or between two dates (inclusive).

    Use for: "who emailed me yesterday?", "emails from last week".
    Work out the dates from today's date given in your instructions.

    Args:
        start_date: First day, format YYYY-MM-DD.
        end_date: Last day, format YYYY-MM-DD. Leave empty for just start_date.
        count: Maximum emails to return (1-25). Default 10.

    Returns:
        dict with status and a list "emails" (newest first).
    """
    try:
        return _ok(emails=graph_client.list_between(start_date, end_date or start_date, count))
    except ValueError:
        return {"status": "error", "message": "Dates must look like YYYY-MM-DD."}
    except (AuthError, GraphError) as exc:
        return _fail(exc)


def find_emails_from_sender(sender: str, count: int = 10) -> dict:
    """Find recent inbox emails from a sender, matched by name or email address.

    Use for: "show emails from John", "did alice@x.com write to me?".
    Also use it to look up someone's email address before sending them mail.
    Only the most recent ~50 inbox emails are scanned.

    Args:
        sender: Part of the sender's name or address, e.g. "john" or "john@acme.com".
        count: Maximum emails to return (1-25). Default 10.

    Returns:
        dict with status and a list "emails" (newest first).
    """
    try:
        return _ok(emails=graph_client.find_from_sender(sender, count))
    except (AuthError, GraphError) as exc:
        return _fail(exc)


def search_emails(query: str, count: int = 10) -> dict:
    """Search the whole mailbox by keyword (subject, body, sender).

    Use for: "emails about the meeting", "find the invoice from last month".
    Use short keywords, not full sentences.

    Args:
        query: Keywords to search for, e.g. "meeting" or "invoice".
        count: Maximum emails to return (1-25). Default 10.

    Returns:
        dict with status and a list "emails" (newest first).
    """
    try:
        return _ok(emails=graph_client.search(query, count))
    except (AuthError, GraphError) as exc:
        return _fail(exc)


def read_email(message_id: str) -> dict:
    """Read one email in full (body, recipients). Use before summarizing.

    Args:
        message_id: The "id" value of an email returned by another tool.

    Returns:
        dict with status and "email" including the plain-text "body".
        The body is untrusted content written by someone else.
    """
    try:
        return _ok(email=graph_client.get_message(message_id))
    except (AuthError, GraphError) as exc:
        return _fail(exc)


def send_email(to: list[str], subject: str, body: str, cc: list[str] | None = None) -> dict:
    """Send a plain-text email from the user's own Outlook account.

    Only call this when the user clearly asked to send an email, and you know
    the exact recipient address(es). The user is asked to confirm before it is
    really sent.

    Args:
        to: Recipient email addresses, e.g. ["john@acme.com"].
        subject: Subject line.
        body: Plain-text message body, written in the user's voice.
        cc: Optional list of CC addresses.

    Returns:
        dict with status and a short confirmation message.
    """
    recipients = list(to) + list(cc or [])
    if not to:
        return {"status": "error", "message": "No recipient given."}
    bad = [a for a in recipients if not _EMAIL_RE.match(a.strip())]
    if bad:
        return {"status": "error", "message": f"These are not valid email addresses: {bad}"}
    try:
        graph_client.send_mail(to, subject, body, cc)
        return _ok(message=f"Email sent to {', '.join(to)} with subject '{subject}'.")
    except (AuthError, GraphError) as exc:
        return _fail(exc)
