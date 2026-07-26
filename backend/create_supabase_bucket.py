import os
from supabase import create_client
from dotenv import load_dotenv

# Load env variables from .env.local
env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env.local')
load_dotenv(env_path)

url = os.getenv('NEXT_PUBLIC_SUPABASE_URL')
key = os.getenv('NEXT_PUBLIC_SUPABASE_ANON_KEY')
sp = create_client(url, key)

print("Attempting to create bucket 'templates'...")
try:
    res = sp.storage.create_bucket("templates", options={"public": True})
    print("Bucket created:", res)
except Exception as e:
    print("Error creating bucket:", e)

# Re-test uploading file to bucket 'templates'
file_path = r"C:\Users\ROHITH P\Downloads\WORKSHOP.pptx"
if os.path.exists(file_path):
    with open(file_path, "rb") as f:
        file_bytes = f.read()
    
    try:
        res = sp.storage.from_("templates").upload("ML/WORKSHOP.pptx", file_bytes, {"upsert": "true", "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"})
        print("Upload result:", res)
    except Exception as e:
        print("Upload error:", e)
