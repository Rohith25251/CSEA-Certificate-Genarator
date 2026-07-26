import os
from dotenv import load_dotenv
from supabase import create_client

# Load env variables from .env.local
env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env.local')
load_dotenv(env_path)

url = os.getenv('NEXT_PUBLIC_SUPABASE_URL')
key = os.getenv('NEXT_PUBLIC_SUPABASE_ANON_KEY')

sp = create_client(url, key)

print("--- Testing 'events' table select ---")
try:
    res = sp.table('events').select('*').execute()
    print("Events count:", len(res.data))
except Exception as e:
    print("Events query notice:", e)

print("\n--- Testing 'students' table select ---")
try:
    res = sp.table('students').select('*').execute()
    print("Students count:", len(res.data))
    if res.data:
        print("Student record keys:", list(res.data[0].keys()))
except Exception as e:
    print("Students query notice:", e)

print("\n--- Testing 'certificates' table select ---")
try:
    res = sp.table('certificates').select('*').execute()
    print("Certificates count:", len(res.data))
    if res.data:
        print("Certificate record keys:", list(res.data[0].keys()))
except Exception as e:
    print("Certificates query notice:", e)
