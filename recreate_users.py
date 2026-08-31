import os
import secrets
import string
from supabase import create_client

url = 'https://dljknrumyawpvxdvjazn.supabase.co'
key = os.environ.get('SUPABASE_SERVICE_KEY')
sb = create_client(url, key)

USERS = ['nadhira@wearcheckrs.com', 'andrew@wearcheckrs.com', 'megan@wearcheckrs.com']

def make_password():
    chars = string.ascii_letters + string.digits
    return ''.join(secrets.choice(chars) for _ in range(10))

users = sb.auth.admin.list_users()

print()
for email in USERS:
    user = next((u for u in users if u.email == email), None)

    # Delete existing account
    if user:
        sb.auth.admin.delete_user(user.id)
        print(f'Deleted old account: {email}')

    # Recreate fresh with password
    pw = make_password()
    result = sb.auth.admin.create_user({
        'email': email,
        'password': pw,
        'email_confirm': True,
    })
    print(f'Created : {result.user.email}')
    print(f'Password: {pw}')
    print()

print('Both accounts recreated. Share the passwords above.')
