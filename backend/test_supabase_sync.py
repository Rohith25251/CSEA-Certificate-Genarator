import os
from supabase import create_client
from dotenv import load_dotenv

# Load env variables from .env.local
env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env.local')
load_dotenv(env_path)

url = os.getenv('NEXT_PUBLIC_SUPABASE_URL')
key = os.getenv('NEXT_PUBLIC_SUPABASE_ANON_KEY')
sp = create_client(url, key)

print("--- Querying 'students' table ---")
try:
    res = sp.table('students').select('*').execute()
    print("Students count:", len(res.data))
    print("Sample student record:", res.data[:2])
except Exception as e:
    print("Error querying students:", e)

print("\n--- Querying 'certificates' table ---")
try:
    res_cert = sp.table('certificates').select('*').execute()
    print("Certificates count:", len(res_cert.data))
    print("Sample cert record:", res_cert.data[:2])
except Exception as e:
    print("Error querying certificates:", e)
