import os
import re
import io
import uuid
import datetime
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
from pydantic import BaseModel
from typing import List, Dict, Any, Optional

from supabase_client import (
    get_students_from_db,
    get_certificates_from_db,
    upsert_event_to_db,
    upsert_students_to_db,
    upsert_certificates_to_db,
    supabase
)
from generator import (
    extract_placeholders_from_pptx,
    parse_excel_dataframe,
    generate_single_native_pdf,
    build_dynamic_replacements,
    pdf_to_base64_png,
    get_active_pptx_template_path,
    SAVED_PPTX_PATH,
    DEFAULT_CERTIFICATE_HTML
)
from mailer import send_certificate_email

app = FastAPI(
    title="CSEA Certificate Generator Python API",
    description="Native PowerPoint PDF Generation Engine & Supabase DB Sync API",
    version="1.0.0"
)

# Enable CORS for Next.js frontend (localhost:3000)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class CertificateRunRequest(BaseModel):
    batch_id: Optional[str] = "Batch_WRK_2026"
    event_id: str
    event_name: str
    event_category: Optional[str] = "Workshop"
    event_date: str
    issue_date: str
    html_template: Optional[str] = None
    rows: List[Dict[str, Any]]

class SinglePreviewRequest(BaseModel):
    event_id: Optional[str] = "evt-ml-workshop-2026"
    event_name: str
    event_category: Optional[str] = "Workshop"
    event_date: str
    issue_date: str
    row: Dict[str, Any]

OUTPUT_CERTS_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), 'generated_pdfs'))
TEMPLATES_BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), 'event_templates'))
os.makedirs(OUTPUT_CERTS_DIR, exist_ok=True)
os.makedirs(TEMPLATES_BASE_DIR, exist_ok=True)

def extract_roll_no_from_row(row: dict, idx: int = 0) -> str:
    """Extracts student roll number / register number from Excel row handling all column alias variations."""
    if not row:
        return f"25CSR{100 + idx}"
    keys = [
        'RollNumber', 'Roll Number ', 'Roll Number', 'Roll No', 'RollNo', 
        'Roll_Number', 'Roll_No', 'Register No', 'Register_No', 'Reg No', 
        'Reg_No', 'Student ID', 'StudentID', 'roll_no', 'register_no'
    ]
    for k in keys:
        if row.get(k) and str(row[k]).strip():
            return str(row[k]).strip()
    
    for k, v in row.items():
        if k and v and re.search(r'(roll|reg|student.*id)', str(k), re.I) and str(v).strip():
            return str(v).strip()

    return f"25CSR{100 + idx}"

def extract_email_from_row(r: dict, roll_no: str) -> str:
    """Extracts email address from Excel row dynamically searching all email/mail header variations."""
    if not r:
        return f"{roll_no.lower()}@kongu.edu"
    
    # 1. Search for any key containing 'mail' or 'email'
    for k, v in r.items():
        if not k or not v:
            continue
        clean_k = str(k).strip().lower()
        clean_v = str(v).strip()
        if ('mail' in clean_k or 'email' in clean_k) and '@' in clean_v:
            return clean_v

    # 2. Fallback check for any string value containing '@'
    for k, v in r.items():
        clean_v = str(v).strip()
        if '@' in clean_v and '.' in clean_v:
            return clean_v

    return f"{roll_no.lower()}@kongu.edu"

@app.get("/")
def read_root():
    return {
        "status": "online",
        "service": "CSEA Certificate Generator Native Python API",
        "native_pptx_rendering": True,
        "time": datetime.datetime.now().isoformat()
    }

@app.get("/api/students")
def fetch_students():
    """Fetch students directly from Supabase DB."""
    students = get_students_from_db()
    return {"status": "success", "count": len(students), "data": students}

@app.get("/api/certificates")
def fetch_certificates():
    """Fetch certificates directly from Supabase DB."""
    certs = get_certificates_from_db()
    return {"status": "success", "count": len(certs), "data": certs}

@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    """Parses Excel sheet WITHOUT writing students to DB yet (waiting for Stage 02 generation)."""
    contents = await file.read()
    filename = file.filename or "dataset.xlsx"

    parsed = parse_excel_dataframe(contents, filename)
    rows = parsed["rows"]

    return {
        "status": "success",
        "file_name": filename,
        "total_rows": parsed["total_rows"],
        "headers": parsed["headers"],
        "rows": rows
    }

@app.post("/api/upload-pptx")
async def upload_pptx(file: UploadFile = File(...)):
    """Parses PPTX PowerPoint template and saves for native PowerPoint PDF rendering."""
    contents = await file.read()
    with open(SAVED_PPTX_PATH, "wb") as f:
        f.write(contents)

    result = extract_placeholders_from_pptx(contents)
    return {
        "status": "success",
        "file_name": file.filename,
        "placeholders": result["placeholders"],
        "template_saved": True
    }

@app.post("/api/register-event")
async def register_event(
    event_id: str = Form(...),
    event_name: str = Form(...),
    event_category: str = Form("Workshop"),
    event_date: str = Form("2026-07-25"),
    file: Optional[UploadFile] = File(None)
):
    """Registers ONLY the event in Supabase events table and saves template in event_id subfolder & Supabase Storage."""
    # 1. Save Event to Supabase DB
    upsert_event_to_db({
        "event_id": event_id,
        "event_name": event_name,
        "event_category": event_category,
        "event_date": event_date
    })

    # 2. Save Template File into Event-ID subfolder and upload to Supabase Storage
    template_saved_path = None
    if file:
        event_folder = os.path.join(TEMPLATES_BASE_DIR, event_id)
        os.makedirs(event_folder, exist_ok=True)
        template_saved_path = os.path.join(event_folder, file.filename or "template.pptx")
        contents = await file.read()
        with open(template_saved_path, "wb") as f:
            f.write(contents)

        with open(SAVED_PPTX_PATH, "wb") as f:
            f.write(contents)

        # Upload directly to Supabase Storage bucket 'templates'
        if supabase:
            try:
                storage_path = f"{event_id}/{file.filename or 'template.pptx'}"
                supabase.storage.from_("templates").upload(
                    storage_path,
                    contents,
                    {"upsert": "true", "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"}
                )
                print(f"[Supabase Storage] Successfully uploaded template to {storage_path}")
            except Exception as st_err:
                print(f"[Supabase Storage] Notice: {st_err}")

    return {
        "status": "success",
        "event_id": event_id,
        "event_name": event_name,
        "template_folder": os.path.join("event_templates", event_id),
        "template_saved": bool(template_saved_path)
    }

@app.post("/api/generate-certificates")
def generate_certificates(req: CertificateRunRequest):
    """
    Called ONLY when clicking 'Generate Certificates' in Stage 02:
    Compiles PDF certificates, upserts Students & Certificates into Supabase DB.
    """
    # 1. Upsert Event to Supabase events table
    upsert_event_to_db({
        "event_id": req.event_id,
        "event_name": req.event_name,
        "event_category": req.event_category or "Workshop",
        "event_date": req.event_date
    })

    rows = req.rows
    students_to_db = []
    certs_to_db = []
    generated_certs = []

    for idx, r in enumerate(rows):
        roll_no = extract_roll_no_from_row(r, idx)
        name = str(r.get('Name', r.get('name', 'Participant'))).strip()
        email = extract_email_from_row(r, roll_no)
        phone = str(r.get('Mobile number ', r.get('phone', ''))).strip()
        section = str(r.get('Section', r.get('section', ''))).strip()
        college = str(r.get('College Name', r.get('College', 'Kongu Engineering College'))).strip() or "Kongu Engineering College"
        year = str(r.get('Year of Study', r.get('Year', r.get('year', '')))).strip()

        students_to_db.append({
            "register_no": roll_no,
            "name": name,
            "email": email,
            "phone": phone,
            "department": "Computer Science and Engineering",
            "year_of_study": year,
            "section": section,
            "college_name": college
        })

        cert_code = f"CSEA-2026-WRK-{idx+1:03d}{uuid.uuid4().hex[:4].upper()}"
        issue_date = req.issue_date or datetime.date.today().isoformat()
        event_name = req.event_name or "DATASET TO DECISION Workshop"
        event_date = req.event_date or issue_date

        replacements = build_dynamic_replacements(r, {
            "certificate_id": cert_code,
            "student_id": roll_no,
            "Name": name,
            "Name ": name,
            "name": name,
            "Roll Number": roll_no,
            "Roll Number ": roll_no,
            "roll_no": roll_no,
            "Register No": roll_no,
            "issue_date": issue_date,
            "event_date": event_date,
            "event_name": event_name,
            "Event Name": event_name,
            "college_name": college,
            "Date": issue_date,
            "Date ": issue_date
        })

        safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
        cert_pdf_name = f"Cert_{cert_code}_{safe_name}.pdf"
        cert_pdf_path = os.path.join(OUTPUT_CERTS_DIR, cert_pdf_name)

        active_template = get_active_pptx_template_path(req.event_id)
        has_native = generate_single_native_pdf(active_template, replacements, cert_pdf_path)

        cert_record = {
            "student_id": roll_no,
            "event_id": req.event_id,
            "student_name": name,
            "issue_date": issue_date,
            "email_status": "pending",
            "created_at": datetime.datetime.now().isoformat()
        }

        certs_to_db.append(cert_record)
        generated_certs.append({
            "certificateCode": cert_code,
            "studentName": name,
            "studentId": roll_no,
            "studentEmail": email,
            "issueDate": issue_date,
            "pdfFilename": cert_pdf_name
        })

    # 2. Upsert Students to Supabase
    inserted_students = upsert_students_to_db(students_to_db)

    # 3. Upsert Certificates to Supabase
    upsert_certificates_to_db(certs_to_db, inserted_students)

    return {
        "status": "success",
        "batch_id": req.batch_id,
        "event_id": req.event_id,
        "count": len(generated_certs),
        "certificates": generated_certs
    }

@app.get("/api/download-pdf/{filename}")
def download_pdf(filename: str, mode: Optional[str] = "inline"):
    """
    Dynamically generates and serves PDF files directly from the latest Supabase DB values
    and the latest PPTX template from Supabase Storage bucket.
    """
    clean_target = filename.replace('.pdf', '').replace('Cert_', '').strip()
    target_parts = [p.lower() for p in clean_target.split('_') if len(p) >= 2 and not p.startswith('CSEA')]

    # 1. Query Supabase DB for matching student & certificate record
    matched_record = None
    if supabase:
        try:
            res = supabase.table('certificates').select('*, events(event_name), students(*)').execute()
            if res.data:
                for item in res.data:
                    st = item.get('students') or {}
                    st_name = str(st.get('name') or item.get('student_name') or '').strip().lower()
                    st_reg = str(st.get('register_no') or item.get('student_id') or '').strip().lower()
                    st_email = str(st.get('email') or '').strip().lower()
                    c_id = str(item.get('id') or '').strip().lower()

                    if (clean_target.lower() in (c_id, st_name, st_reg, st_email) or
                        any(p in st_name or p in st_reg for p in target_parts if len(p) > 2)):
                        matched_record = item
                        break
        except Exception as e:
            print(f"[Supabase] Notice querying DB for dynamic PDF rendering: {e}")

    if matched_record:
        st = matched_record.get('students') or {}
        ev = matched_record.get('events') or {}

        event_id = matched_record.get('event_id') or ''
        event_name = ev.get('event_name') or matched_record.get('event_name') or 'Workshop'
        student_name = st.get('name') or matched_record.get('student_name') or 'Participant'
        roll_no = st.get('register_no') or matched_record.get('student_id') or ''
        issue_date = matched_record.get('issue_date') or datetime.date.today().isoformat()
        college = st.get('college_name') or 'Kongu Engineering College'

        # Fetch latest PPTX template path from Supabase Storage bucket
        active_template = get_active_pptx_template_path(event_id)

        replacements = build_dynamic_replacements({
            "Name": student_name,
            "Name ": student_name,
            "name": student_name,
            "RollNumber": roll_no,
            "Roll Number": roll_no,
            "Roll Number ": roll_no,
            "Register No": roll_no,
            "Date": issue_date,
            "Date ": issue_date,
            "event_name": event_name,
            "Event Name": event_name,
            "college_name": college
        })

        safe_name = "".join(c for c in student_name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
        pdf_path = os.path.join(OUTPUT_CERTS_DIR, f"Cert_{safe_name}.pdf")

        generate_single_native_pdf(active_template, replacements, pdf_path)

        if os.path.exists(pdf_path):
            disp = "attachment" if mode == "attachment" else "inline"
            headers = {
                "Content-Disposition": f'{disp}; filename="Cert_{safe_name}.pdf"',
                "Cache-Control": "no-cache, no-store, must-revalidate",
                "Pragma": "no-cache",
                "Expires": "0"
            }
            return FileResponse(pdf_path, media_type='application/pdf', headers=headers)

    # 2. Fallback to cached disk PDFs
    pdf_path = os.path.join(OUTPUT_CERTS_DIR, filename)
    if not os.path.exists(pdf_path):
        all_pdfs = [f for f in os.listdir(OUTPUT_CERTS_DIR) if f.endswith('.pdf')]
        matching_pdfs = []
        for f in all_pdfs:
            f_lower = f.lower()
            if any(part in f_lower for part in target_parts if len(part) > 2):
                matching_pdfs.append(f)

        if matching_pdfs:
            matching_pdfs.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_CERTS_DIR, x)), reverse=True)
            filename = matching_pdfs[0]
            pdf_path = os.path.join(OUTPUT_CERTS_DIR, filename)
        elif all_pdfs:
            all_pdfs.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_CERTS_DIR, x)), reverse=True)
            filename = all_pdfs[0]
            pdf_path = os.path.join(OUTPUT_CERTS_DIR, filename)

    if os.path.exists(pdf_path):
        disp = "attachment" if mode == "attachment" else "inline"
        headers = {
            "Content-Disposition": f'{disp}; filename="{filename}"',
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0"
        }
        return FileResponse(pdf_path, media_type='application/pdf', headers=headers)

    raise HTTPException(status_code=404, detail="PDF certificate file not found")

class SendEmailRequest(BaseModel):
    student_email: str
    student_name: str
    certificate_code: Optional[str] = "CERT-001"
    certificate_id: Optional[str] = None
    pdf_filename: Optional[str] = None
    event_name: Optional[str] = "Workshop"
    event_date: Optional[str] = "2026-07-25"
    logo_img_url: Optional[str] = None
    hero_img_url: Optional[str] = None

@app.post("/api/send-email-single")
def send_single_email(req: SendEmailRequest):
    """Dispatches a single certificate email using the latest Supabase DB values and Storage template."""
    pdf_path = None
    target_email = req.student_email.strip().lower()
    student_name = req.student_name
    event_name = req.event_name or "Workshop"
    event_date = req.event_date or "2026-07-25"

    # Query DB for latest student & certificate details
    if supabase:
        try:
            res = supabase.table('certificates').select('*, events(event_name), students(*)').execute()
            if res.data:
                for item in res.data:
                    st = item.get('students') or {}
                    st_email = str(st.get('email') or '').strip().lower()
                    if st_email == target_email or target_email in st_email:
                        ev = item.get('events') or {}
                        student_name = st.get('name') or item.get('student_name') or req.student_name
                        roll_no = st.get('register_no') or item.get('student_id') or ''
                        event_id = item.get('event_id') or ''
                        event_name = ev.get('event_name') or item.get('event_name') or req.event_name or 'Workshop'
                        issue_date = item.get('issue_date') or req.event_date or datetime.date.today().isoformat()
                        college = st.get('college_name') or 'Kongu Engineering College'

                        active_template = get_active_pptx_template_path(event_id)
                        replacements = build_dynamic_replacements({
                            "Name": student_name,
                            "Name ": student_name,
                            "name": student_name,
                            "RollNumber": roll_no,
                            "Roll Number": roll_no,
                            "Roll Number ": roll_no,
                            "Register No": roll_no,
                            "Date": issue_date,
                            "Date ": issue_date,
                            "event_name": event_name,
                            "Event Name": event_name,
                            "college_name": college
                        })
                        safe_name = "".join(c for c in student_name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
                        pdf_path = os.path.join(OUTPUT_CERTS_DIR, f"Cert_{safe_name}.pdf")
                        generate_single_native_pdf(active_template, replacements, pdf_path)
                        break
        except Exception as e:
            print(f"[Mailer] Notice generating dynamic PDF for email: {e}")

    if not pdf_path or not os.path.exists(pdf_path):
        if req.pdf_filename:
            pdf_path = os.path.join(OUTPUT_CERTS_DIR, req.pdf_filename)
            if not os.path.exists(pdf_path):
                safe_name = "".join(c for c in student_name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
                for f in os.listdir(OUTPUT_CERTS_DIR):
                    if safe_name.lower() in f.lower() and f.endswith('.pdf'):
                        pdf_path = os.path.join(OUTPUT_CERTS_DIR, f)
                        break

    res = send_certificate_email(
        recipient_email=target_email,
        student_name=student_name,
        cert_code=req.certificate_code or "CERT-001",
        pdf_path=pdf_path or "",
        event_name=event_name,
        event_date=event_date,
        custom_hero_url=req.hero_img_url,
        custom_logo_url=req.logo_img_url,
        cert_id=req.certificate_id
    )
    return res

@app.post("/api/send-email-batch")
def send_batch_emails(items: List[SendEmailRequest]):
    """Batch sends certificate emails with status report."""
    results = []
    for item in items:
        pdf_path = None
        if item.pdf_filename:
            pdf_path = os.path.join(OUTPUT_CERTS_DIR, item.pdf_filename)
            if not os.path.exists(pdf_path):
                safe_name = "".join(c for c in item.student_name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
                for f in os.listdir(OUTPUT_CERTS_DIR):
                    if safe_name.lower() in f.lower() and f.endswith('.pdf'):
                        pdf_path = os.path.join(OUTPUT_CERTS_DIR, f)
                        break

        res = send_certificate_email(
            recipient_email=item.student_email,
            student_name=item.student_name,
            cert_code=item.certificate_code or "CERT-001",
            pdf_path=pdf_path or "",
            event_name=item.event_name or "Workshop",
            event_date=item.event_date or "2026-07-25",
            custom_hero_url=item.hero_img_url,
            custom_logo_url=item.logo_img_url,
            cert_id=item.certificate_id
        )
        results.append(res)

    sent_count = sum(1 for r in results if r.get("success"))
    return {
        "status": "success",
        "total": len(items),
        "sent": sent_count,
        "results": results
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
