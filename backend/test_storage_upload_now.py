import os
from supabase import create_client
from dotenv import load_dotenv

# Load env variables from .env.local
env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env.local')
load_dotenv(env_path)

url = os.getenv('NEXT_PUBLIC_SUPABASE_URL')
key = os.getenv('NEXT_PUBLIC_SUPABASE_ANON_KEY')
sp = create_client(url, key)

pptx_path = r"C:\Users\ROHITH P\Downloads\WORKSHOP.pptx"
if not os.path.exists(pptx_path):
    pptx_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'uploaded_templates', 'latest_template.pptx'))

print(f"Testing direct Supabase Storage upload: {pptx_path}")
with open(pptx_path, "rb") as f:
    contents = f.read()

# Upload under event_id folder 'ML/WORKSHOP.pptx'
try:
    res = sp.storage.from_("templates").upload(
        "ML/WORKSHOP.pptx",
        contents,
        {"upsert": "true", "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"}
    )
    print("Storage Upload Result:", res)
except Exception as e:
    print("Storage Upload Exception:", e)

# List objects in bucket 'templates'
try:
    files = sp.storage.from_("templates").list("ML")
    print("Files inside ML folder in templates bucket:", files)
except Exception as e:
    print("List error:", e)
