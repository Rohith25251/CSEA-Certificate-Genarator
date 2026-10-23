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
        ev_id = str(event_data['event_id']).strip()
        ev_name = str(event_data.get('event_name') or ev_id).strip()
        ev_date = str(event_data.get('event_date') or '2026-09-21').strip()
        record = {
            "event_id": ev_id,
            "event_date": ev_date
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
        student_dict = {}
        for idx, s in enumerate(students_list):
            reg = str(s.get('register_no', '')).strip()
            name = str(s.get('name', 'Participant')).strip()
            email = str(s.get('email', '')).strip()
            phone = str(s.get('phone', '')).strip()
            dept = str(s.get('department', 'Computer Science and Engineering')).strip()
            year = str(s.get('year_of_study', '')).strip()
            sec = str(s.get('section', '')).strip()
            college = str(s.get('college_name') or 'Kongu Engineering College').strip()

            if not reg:
                # If roll number is missing, create a deterministic temp fallback
                reg = f"temp{idx + 1}"

            reg_key = reg.lower()
            if reg_key in student_dict:
                existing = student_dict[reg_key]
                # If this entry is for the same student (same name/email or one is empty), enrich fields
                same_person = (
                    not existing['email'] or not email or existing['email'].lower() == email.lower()
                    or existing['name'].lower() == name.lower()
                )
                if same_person:
                    if email and not existing['email']:
                        existing['email'] = email
                    if phone and not existing['phone']:
                        existing['phone'] = phone
                    if year and not existing['year_of_study']:
                        existing['year_of_study'] = year
                    if sec and not existing['section']:
                        existing['section'] = sec
                    if name != 'Participant' and existing['name'] == 'Participant':
                        existing['name'] = name
                else:
                    # Different person mistakenly assigned the same roll number in Excel
                    reg = f"{reg}-{idx+1}"
                    reg_key = reg.lower()
                    student_dict[reg_key] = {
                        "register_no": reg,
                        "name": name,
                        "email": email,
                        "phone": phone,
                        "department": dept,
                        "year_of_study": year,
                        "section": sec,
                        "college_name": college
                    }
            else:
                student_dict[reg_key] = {
                    "register_no": reg,
                    "name": name,
                    "email": email,
                    "phone": phone,
                    "department": dept,
                    "year_of_study": year,
                    "section": sec,
                    "college_name": college
                }

        clean_students = list(student_dict.values())

        # Batch upsert in chunks to ensure no request exceeds payload limits
        chunk_size = 50
        upserted_data = []
        for i in range(0, len(clean_students), chunk_size):
            chunk = clean_students[i:i + chunk_size]
            try:
                res = supabase.table('students').upsert(chunk, on_conflict='register_no').execute()
                if res.data:
                    upserted_data.extend(res.data)
            except Exception as chunk_err:
                print(f"[Supabase] Chunk upsert warning: {chunk_err}. Retrying row-by-row...")
                for item in chunk:
                    try:
                        single_res = supabase.table('students').upsert([item], on_conflict='register_no').execute()
                        if single_res.data:
                            upserted_data.extend(single_res.data)
                    except Exception as single_err:
                        print(f"[Supabase] Single student upsert error for {item.get('register_no')}: {single_err}")

        print(f"[Supabase] Successfully upserted {len(upserted_data)} students!")
        return upserted_data
    except Exception as e:
        print(f"[Supabase] Failed to upsert students: {e}")
        return []

def upsert_certificates_to_db(certs_list, inserted_students=None):
    if not supabase or not certs_list:
        return False
    try:
        # 1. Pre-register distinct events in certs_list to 'events' table to satisfy FK constraint
        distinct_events = {}
        for c in certs_list:
            ev_id = str(c.get('event_id') or '').strip()
            ev_date = str(c.get('event_date') or '').strip()
            if ev_id and ev_id not in distinct_events:
                distinct_events[ev_id] = ev_date

        for ev_id, ev_date in distinct_events.items():
            upsert_event_to_db({
                "event_id": ev_id,
                "event_date": ev_date or "2026-09-21"
            })

        # 2. Build multi-tier student lookup dictionary
        student_map = {}
        email_map = {}
        name_map = {}

        if inserted_students:
            for s in inserted_students:
                reg = str(s.get('register_no', '')).strip().lower()
                em = str(s.get('email', '')).strip().lower()
                nm = str(s.get('name', '')).strip().lower()
                sid = s.get('id')
                if sid:
                    if reg:
                        student_map[reg] = sid
                    if em:
                        email_map[em] = sid
                    if nm:
                        name_map[nm] = sid

        db_students = get_students_from_db()
        for s in db_students:
            reg = str(s.get('register_no', '')).strip().lower()
            em = str(s.get('email', '')).strip().lower()
            nm = str(s.get('name', '')).strip().lower()
            sid = s.get('id')
            if sid:
                if reg:
                    student_map[reg] = sid
                if em:
                    email_map[em] = sid
                if nm:
                    name_map[nm] = sid

        clean_certs = []
        for c in certs_list:
            st_id = str(c.get('student_id', '')).strip()

            valid_uuid = None
            if UUID_REGEX.match(st_id):
                valid_uuid = st_id
            elif st_id.lower() in student_map:
                valid_uuid = student_map[st_id.lower()]

            if not valid_uuid:
                s_email = str(c.get('student_email', '')).strip().lower()
                if s_email in email_map:
                    valid_uuid = email_map[s_email]

            if not valid_uuid:
                s_name = str(c.get('student_name', '')).strip().lower()
                if s_name in name_map:
                    valid_uuid = name_map[s_name]

            clean_certs.append({
                "student_id": valid_uuid,
                "event_id": str(c.get('event_id') or '').strip() or 'Workshop',
                "student_name": c.get('student_name', 'Participant'),
                "student_email": c.get('student_email', ''),
                "event_name": c.get('event_name', '') or '',
                "title": c.get('title', '') or '',
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
    
    import re

    def _sanitize(name: str) -> str:
        if not name:
            return "default"
        s = re.sub(r'\s*[/\\:]+\s*', '_', str(name))
        s = re.sub(r'[<>\"|?*\x00-\x1f]', '_', s)
        s = s.strip().strip('.').strip()
        return s or "default"

    target_folders = []
    if event_id and str(event_id).strip():
        clean_eid = str(event_id).strip()
        safe_eid = _sanitize(clean_eid)
        target_folders.append(safe_eid)
        if clean_eid != safe_eid:
            target_folders.append(clean_eid)
    else:
        target_folders.append("")  # Root folder ONLY if event_id is not specified

    is_vercel = os.environ.get('VERCEL') is not None or os.environ.get('AWS_LAMBDA_FUNCTION_NAME') is not None
    if is_vercel:
        save_dir = os.path.join('/tmp', 'uploaded_templates')
    else:
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
                    safe_slug = _sanitize(folder) if folder else "latest"
                    save_name = f"{safe_slug}_template.pptx"
                    save_path = os.path.join(save_dir, save_name)
                    with open(save_path, 'wb') as f:
                        f.write(content)
                    print(f"[Supabase Storage] Downloaded latest template '{storage_path}' to {save_path}")
                    return save_path
        except Exception as e:
            print(f"[Supabase Storage] Notice checking template folder '{folder}': {e}")

    return ""
