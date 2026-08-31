"""
One-time script to create WearCheck ARC user accounts with temporary passwords.
Share each password with the user and ask them to change it after first login.

Usage (PowerShell):
  $env:SUPABASE_SERVICE_KEY="eyJ..."
  python create_users.py
"""

import os
import sys
import secrets
import string
from supabase import create_client

USERS = [
    "nadhira@wearcheckrs.com",
    "andrew@wearcheckrs.com",
    "megan@wearcheckrs.com",
]

def make_password():
    alphabet = string.ascii_letters + string.digits + "!@#$"
    return ''.join(secrets.choice(alphabet) for _ in range(12))

def main():
    url = os.environ.get("SUPABASE_URL", "https://dljknrumyawpvxdvjazn.supabase.co")
    key = os.environ.get("SUPABASE_SERVICE_KEY")

    if not key:
        print("\nERROR: SUPABASE_SERVICE_KEY is not set.")
        print("Set it first, then re-run:\n")
        print('  $env:SUPABASE_SERVICE_KEY="eyJ..."')
        print("  python create_users.py\n")
        sys.exit(1)

    sb = create_client(url, key)

    print("\n" + "="*60)
    for email in USERS:
        temp_password = make_password()
        try:
            sb.auth.admin.create_user({
                "email": email,
                "password": temp_password,
                "email_confirm": True,
            })
            print(f"  Email   : {email}")
            print(f"  Password: {temp_password}")
            print(f"  Status  : Created OK")
        except Exception as e:
            msg = str(e)
            if "already been registered" in msg or "already exists" in msg:
                # Update the password instead
                try:
                    users = sb.auth.admin.list_users()
                    user = next((u for u in users if u.email == email), None)
                    if user:
                        sb.auth.admin.update_user_by_id(user.id, {"password": temp_password})
                        print(f"  Email   : {email}")
                        print(f"  Password: {temp_password}")
                        print(f"  Status  : Already existed — password updated")
                    else:
                        print(f"  Email   : {email}")
                        print(f"  Status  : Already exists, could not find to update")
                except Exception as e2:
                    print(f"  Email   : {email}")
                    print(f"  Status  : FAILED — {e2}")
            else:
                print(f"  Email   : {email}")
                print(f"  Status  : FAILED — {msg}")
        print()

    print("="*60)
    print("Share each password directly with the user (e.g. via Teams/WhatsApp).")
    print("Ask them to change it after signing in via 'Forgot password?'.\n")

if __name__ == "__main__":
    main()
