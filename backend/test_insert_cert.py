import datetime
from supabase import create_client
import os
from dotenv import load_dotenv

# Load env variables from .env.local
env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env.local')
load_dotenv(env_path)

url = os.getenv('NEXT_PUBLIC_SUPABASE_URL')
key = os.getenv('NEXT_PUBLIC_SUPABASE_ANON_KEY')
sp = create_client(url, key)

# Get student record UUID
stu_res = sp.table('students').select('*').eq('register_no', '25CSR220').execute()
stu_uuid = stu_res.data[0]['id'] if stu_res.data else None
print("Student UUID:", stu_uuid)

sample_cert = {
    "certificate_code": "CSEA-2026-WRK-001",
    "student_name": "Preethika sri K",
    "student_email": "25csr220@kongu.edu",
    "student_id": stu_uuid,
    "issue_date": "2026-07-25",
    "email_status": "delivered",
    "email_sent_at": datetime.datetime.now().isoformat(),
    "custom_fields": {
        "Name": "Preethika sri K",
        "Roll Number": "25CSR220",
        "Section": "D",
        "event_name": "DATASET TO DECISION Workshop"
    }
}

print("Attempting insert into 'certificates'...")
try:
    res = sp.table('certificates').upsert([sample_cert], on_conflict='certificate_code').execute()
    print("Insert success! Inserted rows:", len(res.data))
    print("Inserted data:", res.data)
except Exception as e:
    print("Insert error:", e)
