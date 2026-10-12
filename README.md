# Outlook Email Assistant (Google ADK + Microsoft Graph + Gemini)

A local AI agent that connects to your Microsoft Outlook inbox. You chat with it in plain English and it can:

- **Read** emails (sender, subject, body, date)
- **Answer questions** about them ("Who emailed me about the meeting yesterday?")
- **Act**: compose and **send** emails (with a confirmation step before anything is sent)

| Piece | Choice |
|---|---|
| Agent framework | Google Agent Development Kit (`google-adk` 2.11.0) |
| LLM | Gemini API, model `gemini-3.5-flash-lite` (stable; switch to `gemini-3.8-flash` via `GEMINI_MODEL` if you want a stronger one) |
| Mail | Microsoft Graph API v1.0 (Outlook) |
| Auth | OAuth 2.0 via MSAL (`msal` 1.39.0), delegated permissions, token cached locally |
| Language | Python 3.10+ (tested on 3.13) |

---

## 1. Setup

```bash
git clone https://github.com/DivyaMusahib/email-assistant-adk.git
cd email-assistant-adk

python3 -m venv .venv
source .venv/bin/activate            # Windows PowerShell: .venv\Scripts\Activate.ps1
pip install -r requirements.txt

cp .env.example .env                 # Windows: copy .env.example .env
```

Now fill `.env` (two secrets needed):

### 1a. Gemini API key -> `GOOGLE_API_KEY`
1. Open <https://aistudio.google.com/apikey> and sign in with a Google account.
2. Click **Create API key**, copy it.
3. Paste it into `.env` as `GOOGLE_API_KEY=...`

### 1b. Register the app in Microsoft Entra -> `MS_CLIENT_ID`
(Free. Works with a personal outlook.com/hotmail.com account or a work account.)

1. Go to <https://entra.microsoft.com> -> **Entra ID** -> **App registrations** -> **New registration**.
   (If you only have a personal account and the portal complains about no tenant, use <https://portal.azure.com> -> "App registrations"; a free Azure sign-up creates a default directory.)
2. **Name:** `email-assistant-adk`.
3. **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*.
4. **Redirect URI:** platform **Public client/native (mobile & desktop)**, value `http://localhost`.
5. Click **Register**. On the Overview page copy **Application (client) ID** -> `.env` as `MS_CLIENT_ID=...`
6. Left menu **Authentication** -> scroll to **Advanced settings** -> **Allow public client flows = Yes** -> Save. (Needed for the device-code option.)
7. Left menu **API permissions** -> **Add a permission** -> **Microsoft Graph** -> **Delegated permissions** -> tick **Mail.Read**, **Mail.Send** (and **User.Read** if not already there) -> **Add permissions**.
   No admin consent is needed for personal accounts. For a work account, either consent at first login or ask your admin.
8. Keep `MS_TENANT_ID=common` in `.env` (works for both personal and work accounts).

There is **no client secret** to create: this is a local "public client" app.

## 2. Authenticate (do this once)

```bash
python login.py
```

A browser opens -> sign in -> accept the permissions. The terminal prints
`Signed in as <Your Name> <you@outlook.com>`. The login is saved in `.token_cache.json` (git-ignored) and renewed silently, so you do not log in again.

No browser on the machine? Use `python login.py --device-code` and follow the printed instructions.
To forget the login: `python login.py --logout`.

## 3. Run the agent

Terminal chat:
```bash
adk run email_agent
```

Browser chat UI (run from the project root):
```bash
adk web --port 8000
# open http://localhost:8000 and choose "email_agent"
```

When the agent wants to send an email it pauses and asks you to confirm. In `adk run`, type `yes` to send, anything else cancels. In `adk web`, use the approve/reject buttons.

## 4. Example prompts for testing

Retrieval
- `Who sent the latest email?`
- `Show my last 5 emails.`
- `Show emails from John.`
- `Summarize the latest email.`

Question answering
- `Who emailed me about the meeting yesterday?`
- `Do I have any unread invoices? Who sent them?`
- `What did Priya say in her last email?`

Actions
- `Send an email to John confirming the meeting tomorrow.`
  (the agent looks up John's address from your inbox, drafts it, asks you to confirm)
- `Send an email to someone@example.com with subject "Hello" saying thanks for the update.`

Safety checks (good for a demo)
- Answer `no` at the confirmation -> nothing is sent.
- `Send an email to John` when two different Johns exist -> the agent asks which one.


## 5. Project layout

```
email-assistant-adk/
├── email_agent/            # the ADK agent package (folder name = what you pass to `adk run`)
│   ├── __init__.py         # lets ADK discover the agent
│   ├── agent.py            # root_agent: model + instructions + tools
│   ├── tools.py            # the 6 tools the LLM can call
│   ├── graph_client.py     # Microsoft Graph HTTP calls (no ADK inside)
│   ├── auth.py             # MSAL login + silent token refresh + cache
│   └── config.py           # all settings read from .env
├── login.py                # one-time sign-in command
├── requirements.txt
├── .env.example
└── README.md
```

## 6. Troubleshooting

| Symptom | Fix |
|---|---|
| `You are not signed in. Run python login.py` | Do step 2. |
| `AADSTS700016 ... application not found` | Wrong `MS_CLIENT_ID` or wrong `MS_TENANT_ID`. |
| `AADSTS50194` / account type error | Set the registration to "any org directory and personal accounts", keep tenant `common`. |
| `403` from Graph | Add `Mail.Read` / `Mail.Send` permissions (step 1b-7), then `python login.py --logout` and log in again. |
| Gemini `404 model not found` | Set `GEMINI_MODEL` in `.env` to a current model, see <https://ai.google.dev/gemini-api/docs/models>. |
| Gemini `429` | Free-tier rate limit; wait a minute. |
| `adk: command not found` | Activate the virtual environment first. |

## Security notes
- `.env` and `.token_cache.json` are git-ignored.
- Only the least-privilege permissions needed are requested (`Mail.Read`, `Mail.Send`, `User.Read`).
- Emails are treated as untrusted input (prompt-injection guard in the agent instructions), and sending always needs your confirmation.
