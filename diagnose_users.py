import os
from supabase import create_client

url = 'https://dljknrumyawpvxdvjazn.supabase.co'
key = os.environ.get('SUPABASE_SERVICE_KEY')
sb = create_client(url, key)

TARGET = ['nadhira@wearcheckrs.com', 'andrew@wearcheckrs.com', 'megan@wearcheckrs.com']
users = sb.auth.admin.list_users()

for u in users:
    if u.email in TARGET:
        providers = [i.provider for i in u.identities] if u.identities else []
        print('Email     :', u.email)
        print('Confirmed :', 'YES' if u.email_confirmed_at else 'NO')
        print('Providers :', providers)
        print()
