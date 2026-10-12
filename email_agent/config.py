"""Central configuration: read everything from environment variables once.

Why a separate file?  Every other module imports its settings from here, so
there is exactly ONE place that touches os.environ.  Easy to explain, easy to test.
"""

import os
from pathlib import Path

from dotenv import load_dotenv

# Project root = the folder that contains `email_agent/`
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Load <project>/.env into os.environ (does nothing if the file is missing).
# ADK's Gemini client later reads GOOGLE_API_KEY from os.environ by itself.
load_dotenv(PROJECT_ROOT / ".env")

# ---- Gemini ---------------------------------------------------------------
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")

# ---- Microsoft Graph --------------------------------------------------------
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
MS_TENANT_ID = os.getenv("MS_TENANT_ID", "common")
MS_AUTHORITY = f"https://login.microsoftonline.com/{MS_TENANT_ID}"

# Delegated permissions = "act as the signed-in user".  Least privilege:
# we only ask for what the three requirements need.
#   User.Read -> who am I (shows the signed-in address)
#   Mail.Read -> read the mailbox
#   Mail.Send -> send mail as me
# (offline_access is added automatically by MSAL so we get a refresh token.)
MS_SCOPES = ["User.Read", "Mail.Read", "Mail.Send"]

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Where the login is remembered between runs. It is git-ignored.
TOKEN_CACHE_PATH = PROJECT_ROOT / ".token_cache.json"

# ---- Behaviour limits -------------------------------------------------------
MAX_EMAILS_PER_CALL = 25      # never pull more than this into the LLM at once
SCAN_WINDOW = 50              # how many recent mails we scan for sender filtering
BODY_CHAR_LIMIT = 4000        # long bodies are cut so we do not waste tokens
