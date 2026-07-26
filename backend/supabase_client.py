import os
import re
from dotenv import load_dotenv
from supabase import create_client, Client

# Load environment variables from .env.local or environment
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env.local'))

SUPABASE_URL = os.getenv('NEXT_PUBLIC_SUPABASE_URL')
SUPABASE_KEY = os.getenv('NEXT_PUBLIC_SUPABASE_ANON_KEY')

supabase: Client = None

if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        print(f"[Supabase] Connection error: {e}")

UUID_REGEX = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.IGNORECASE)

def get_students_from_db():
    if not supabase:
        return []
    try:
        res = supabase.table('students').select('*').order('created_at', desc=True).execute()
        return res.data or []
    except Exception as e:
        print(f"[Supabase] Failed to fetch students: {e}")
        return []

def get_certificates_from_db():
    if not supabase:
        return []
    try:
        res = supabase.table('certificates').select('*').order('created_at', desc=True).execute()
        return res.data or []
    except Exception as e:
        print(f"[Supabase] Failed to fetch certificates: {e}")
        return []

def upsert_event_to_db(event_data):
    if not supabase or not event_data or not event_data.get('event_id'):
        return False
    try:
        record = {
            "event_id": event_data['event_id'],
            "event_name": event_data.get('event_name', 'DATASET TO DECISION Workshop'),
            "event_category": event_data.get('event_category', 'Workshop'),
            "event_date": event_data.get('event_date', '2026-07-25')
        }
        res = supabase.table('events').upsert([record], on_conflict='event_id').execute()
        return True
    except Exception as e:
        print(f"[Supabase] Failed to upsert event: {e}")
        return False

def upsert_students_to_db(students_list):
    if not supabase or not students_list:
        return []
    try:
        clean_students = []
        for s in students_list:
            clean_students.append({
                "register_no": str(s.get('register_no', '')).strip(),
                "name": str(s.get('name', 'Participant')).strip(),
                "email": str(s.get('email', '')).strip(),
                "phone": str(s.get('phone', '')).strip(),
                "department": str(s.get('department', 'Computer Science and Engineering')).strip(),
                "year_of_study": str(s.get('year_of_study', '')).strip(),
                "section": str(s.get('section', '')).strip(),
                "college_name": str(s.get('college_name') or 'Kongu Engineering College').strip()
            })
        res = supabase.table('students').upsert(clean_students, on_conflict='register_no').execute()
        print(f"[Supabase] Successfully upserted {len(res.data or [])} students!")
        return res.data or []
    except Exception as e:
        print(f"[Supabase] Failed to upsert students: {e}")
        return []

def upsert_certificates_to_db(certs_list, inserted_students=None):
    if not supabase or not certs_list:
        return False
    try:
        # Build student lookup dictionary (register_no -> uuid)
        student_map = {}
        
        # 1. Use returned inserted students if provided
        if inserted_students:
            for s in inserted_students:
                reg = str(s.get('register_no', '')).strip().lower()
                sid = s.get('id')
                if reg and sid:
                    student_map[reg] = sid

        # 2. Re-fetch from DB if mapping incomplete
        db_students = get_students_from_db()
        for s in db_students:
            reg = str(s.get('register_no', '')).strip().lower()
            sid = s.get('id')
            if reg and sid:
                student_map[reg] = sid

        clean_certs = []
        for c in certs_list:
            st_id = str(c.get('student_id', '')).strip()

            valid_uuid = None
            if UUID_REGEX.match(st_id):
                valid_uuid = st_id
            elif st_id.lower() in student_map:
                valid_uuid = student_map[st_id.lower()]

            clean_certs.append({
                "student_id": valid_uuid,
                "event_id": c.get('event_id'),
                "student_name": c.get('student_name', 'Participant'),
                "issue_date": c.get('issue_date', '2026-07-25'),
                "email_status": c.get('email_status', 'pending')
            })

        # Delete any pre-existing certificates for matching student_id and event_id to eliminate duplicates
        for c in clean_certs:
            if c.get('student_id') and c.get('event_id'):
                try:
                    supabase.table('certificates').delete().eq('student_id', c['student_id']).eq('event_id', c['event_id']).execute()
                except Exception:
                    pass

        res = supabase.table('certificates').insert(clean_certs).execute()
        print(f"[Supabase] Successfully saved {len(res.data or [])} unique certificates!")
        return True
    except Exception as e:
        print(f"[Supabase] Failed to insert certificates: {e}")
        return False

def download_latest_template_from_supabase(event_id: str = None) -> str:
    """
    Downloads the latest PPTX template from Supabase Storage bucket 'templates'
    (checking '{event_id}/' folder or root 'templates/' folder) and saves it to local disk.
    """
    if not supabase:
        return ""
    
    target_folders = []
    if event_id:
        target_folders.append(str(event_id).strip())
    target_folders.append("")  # Root folder

    save_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'uploaded_templates'))
    os.makedirs(save_dir, exist_ok=True)

    for folder in target_folders:
        try:
            files = supabase.storage.from_('templates').list(folder)
            if not files:
                continue
            pptx_files = [f for f in files if isinstance(f, dict) and str(f.get('name', '')).endswith('.pptx')]
            if pptx_files:
                # Sort by updated_at or created_at to get newest template
                pptx_files.sort(key=lambda x: str(x.get('updated_at') or x.get('created_at') or ''), reverse=True)
                target_file = pptx_files[0]['name']
                storage_path = f"{folder}/{target_file}" if folder else target_file

                content = supabase.storage.from_('templates').download(storage_path)
                if content:
                    save_name = f"{folder}_template.pptx" if folder else "latest_template.pptx"
                    save_path = os.path.join(save_dir, save_name)
                    with open(save_path, 'wb') as f:
                        f.write(content)
                    print(f"[Supabase Storage] Downloaded latest template '{storage_path}' to {save_path}")
                    return save_path
        except Exception as e:
            print(f"[Supabase Storage] Notice checking template folder '{folder}': {e}")

    return ""
