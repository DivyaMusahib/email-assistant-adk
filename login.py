"""One-time Microsoft sign-in.  Usage:

    python login.py                 # opens your browser (recommended)
    python login.py --device-code   # prints a code to type at microsoft.com/devicelogin
    python login.py --logout        # forget the saved login
"""

import argparse

from email_agent import auth, graph_client


def main() -> None:
    parser = argparse.ArgumentParser(description="Sign in to Microsoft for the email agent")
    parser.add_argument("--device-code", action="store_true", help="use device-code flow")
    parser.add_argument("--logout", action="store_true", help="remove the saved login")
    args = parser.parse_args()

    if args.logout:
        print("Logged out." if auth.logout() else "Nothing to log out from.")
        return

    try:
        auth.login(use_device_code=args.device_code)
        me = graph_client.get_profile()  # proves the token really works
    except (auth.AuthError, graph_client.GraphError) as exc:
        raise SystemExit(f"Login problem: {exc}")

    print(f"Signed in as {me['name']} <{me['email']}>")
    print("You can now run:  adk run email_agent")


if __name__ == "__main__":
    main()
