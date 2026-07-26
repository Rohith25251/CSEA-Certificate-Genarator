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

print(f"Uploading template to Supabase Storage: {pptx_path}")

with open(pptx_path, "rb") as f:
    file_bytes = f.read()

try:
    res = sp.storage.from_("templates").upload(
        "ML/WORKSHOP.pptx",
        file_bytes,
        {"upsert": "true", "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"}
    )
    print("SUCCESS! Upload Result:", res)
except Exception as e:
    print("Upload error:", e)
