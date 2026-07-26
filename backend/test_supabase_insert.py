import os
import sys
from dotenv import load_dotenv
from supabase import create_client

env_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../.env.local'))
load_dotenv(env_path)

url = os.getenv('NEXT_PUBLIC_SUPABASE_URL')
key = os.getenv('NEXT_PUBLIC_SUPABASE_ANON_KEY')

print("URL:", url)
print("KEY length:", len(key) if key else 0)

supabase = create_client(url, key)

print("\n--- Testing Events Insert ---")
try:
    res = supabase.table('events').upsert([{
        "event_id": "TEST_EVENT_01",
        "event_name": "Test Event",
        "event_category": "Workshop",
        "event_date": "2026-07-25"
    }], on_conflict='event_id').execute()
    print("Events Insert Result:", res.data)
except Exception as e:
    print("Events Insert Error:", e)

print("\n--- Testing Students Insert ---")
try:
    res = supabase.table('students').upsert([{
        "register_no": "TEST_25CSR999",
        "name": "Test Student",
        "email": "test@kongu.edu",
        "phone": "9876543210",
        "department": "CSE",
        "year_of_study": "II Year",
        "section": "D",
        "college_name": "Kongu Engineering College"
    }], on_conflict='register_no').execute()
    print("Students Insert Result:", res.data)
    student_uuid = res.data[0]['id'] if res.data else None
    print("Student UUID:", student_uuid)
except Exception as e:
    print("Students Insert Error:", e)
    student_uuid = None

print("\n--- Testing Certificates Insert ---")
try:
    res = supabase.table('certificates').insert([{
        "student_id": student_uuid,
        "event_id": "TEST_EVENT_01",
        "student_name": "Test Student",
        "student_email": "test@kongu.edu",
        "issue_date": "2026-07-25",
        "email_status": "sent"
    }]).execute()
    print("Certificates Insert Result:", res.data)
except Exception as e:
    print("Certificates Insert Error:", e)

print("\n--- Testing Fetching ---")
try:
    st_res = supabase.table('students').select('*').execute()
    print("Total Students in DB:", len(st_res.data or []))
    cert_res = supabase.table('certificates').select('*').execute()
    print("Total Certificates in DB:", len(cert_res.data or []))
except Exception as e:
    print("Fetch Error:", e)
