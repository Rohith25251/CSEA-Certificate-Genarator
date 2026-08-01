import os
import io
import re
import uuid
import base64
import tempfile
import datetime
import pandas as pd
from pptx import Presentation


# Attempt win32com and pythoncom imports for native PowerPoint COM rendering on Windows
has_win32com = False
pythoncom_lib = None
try:
    import win32com.client
    import pythoncom
    pythoncom_lib = pythoncom
    has_win32com = True
except Exception as err:
    print(f"[Generator] win32com/pythoncom import notice: {err}")
    has_win32com = False

# Attempt PyMuPDF for PDF page rasterization
has_fitz = False
try:
    import fitz
    has_fitz = True
except Exception:
    has_fitz = False

DEFAULT_CERTIFICATE_HTML = """
<div style="width: 1000px; height: 700px; padding: 45px; background: #ffffff; border: 12px solid #082849; border-radius: 12px; font-family: 'Playfair Display', serif; color: #1e293b; box-sizing: border-box; position: relative;">
  <div style="border: 2px solid #e2e8f0; height: 100%; padding: 40px; border-radius: 8px; text-align: center; display: flex; flex-direction: column; justify-content: space-between; box-sizing: border-box;">
    <div style="margin: 30px 0;">
      <h1 style="font-size: 32px; font-weight: 900; color: #082849;">CERTIFICATE OF PARTICIPATION</h1>
      <h2 style="font-size: 38px; font-weight: 800; color: #0284c7;"><<Name>></h2>
    </div>
  </div>
</div>
"""

UPLOAD_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), 'uploaded_templates'))
os.makedirs(UPLOAD_DIR, exist_ok=True)

SAVED_PPTX_PATH = os.path.join(UPLOAD_DIR, 'latest_template.pptx')
DOWNLOADS_WORKSHOP_PATH = r"C:\Users\ROHITH P\Downloads\WORKSHOP.pptx"

def get_active_pptx_template_path(event_id: str = None) -> str:
    """Returns path to active PPTX template, prioritizing downloading latest template from Supabase Storage bucket."""
    from supabase_client import download_latest_template_from_supabase
    
    if event_id:
        sp_path = download_latest_template_from_supabase(event_id)
        if sp_path and os.path.exists(sp_path):
            return sp_path

    sp_latest = download_latest_template_from_supabase()
    if sp_latest and os.path.exists(sp_latest):
        return sp_latest

    if event_id:
        event_local = os.path.join(UPLOAD_DIR, f"{event_id}_template.pptx")
        if os.path.exists(event_local):
            return event_local

    if os.path.exists(SAVED_PPTX_PATH):
        return SAVED_PPTX_PATH
    if os.path.exists(DOWNLOADS_WORKSHOP_PATH):
        return DOWNLOADS_WORKSHOP_PATH
    return ""

def save_uploaded_pptx_template(pptx_bytes: bytes) -> str:
    """Saves uploaded PPTX bytes to disk as latest_template.pptx."""
    with open(SAVED_PPTX_PATH, 'wb') as f:
        f.write(pptx_bytes)
    return SAVED_PPTX_PATH

def extract_placeholders_from_pptx(pptx_bytes: bytes):
    """Parses PPTX file, extracts text & placeholders, and saves template."""
    save_uploaded_pptx_template(pptx_bytes)
    
    prs = Presentation(io.BytesIO(pptx_bytes))
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    text_runs.append(p.text)
    
    combined_text = "\n".join(text_runs)
    placeholders = list(set(re.findall(r"<<\s*([^>]+?)\s*>>", combined_text)))
    
    return {
        "text": combined_text,
        "placeholders": placeholders,
        "template_saved": True
    }

from pptx.dml.color import RGBColor

def format_date_to_dd_mm_yyyy(val) -> str:
    if val is None:
        return ""
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime("%d\u2011%m\u2011%Y")
    
    val_str = str(val).strip()
    
    # Check if format is yyyy-mm-dd (with optional T... or time)
    match1 = re.match(r'^(\d{4})[-/](\d{2})[-/](\d{2})(?:\s+.*|T.*)?$', val_str)
    if match1:
        y, m, d = match1.groups()
        return f"{d}\u2011{m}\u2011{y}"
        
    # Check if format is already dd-mm-yyyy or dd/mm/yyyy
    match2 = re.match(r'^(\d{2})[-/](\d{2})[-/](\d{4})$', val_str)
    if match2:
        d, m, y = match2.groups()
        return f"{d}\u2011{m}\u2011{y}"
        
    return val_str

def build_dynamic_replacements(row: dict, extra: dict = None) -> dict:
    """
    Builds a universal dynamic replacement dictionary from ANY arbitrary Excel column header,
    handling trailing/leading spaces, casing, underscore/space variations, and common aliases.
    """
    rep = {}
    if extra:
        rep.update(extra)

    # 1. First pass: Add all original Excel column key-value pairs and clean variations
    for raw_k, v in row.items():
        if raw_k is None:
            continue
        val_str = str(v).strip() if v is not None else ""
        clean_k = str(raw_k).strip()

        rep[raw_k] = val_str
        rep[clean_k] = val_str
        rep[clean_k.lower()] = val_str
        rep[clean_k.upper()] = val_str
        rep[clean_k.title()] = val_str
        rep[clean_k.replace('_', ' ')] = val_str
        rep[clean_k.replace(' ', '_')] = val_str
        rep[clean_k.replace('_', ' ').lower()] = val_str

    # 2. Extract standard fields (Name, Roll No, Email, Date, College, Event) if present
    std_name = None
    std_roll = None
    std_email = None

    for k, v in row.items():
        if not k or v is None:
            continue
        ck = str(k).strip().lower().replace('_', ' ')
        cv = str(v).strip()

        if ('name' in ck or 'student' in ck or 'participant' in ck) and not std_name and cv:
            std_name = cv
        if ('roll' in ck or 'register' in ck or 'reg' in ck or 'id' in ck or 'number' in ck) and not std_roll and cv:
            std_roll = cv
        if ('mail' in ck or 'email' in ck) and '@' in cv:
            std_email = cv

    # 3. Register standard aliases for PowerPoint templates with varied token naming
    if std_name:
        for alias in ["Name", "Name ", "name", "Full Name", "Full Name ", "Student Name", "Participant Name", "PARTICIPANT_NAME"]:
            if alias not in rep:
                rep[alias] = std_name

    if std_roll:
        for alias in [
            "RollNumber", "Roll Number", "Roll Number ", "Roll No", "Roll No ", "Roll_Number", "Roll_No", 
            "RollNo", "Register No", "Register No ", "RegisterNo", "Register_No", "registration_no", 
            "Registration Number", "REGISTRATION_NO", "Reg No", "Reg_No", "Student ID", "StudentID"
        ]:
            rep[alias] = std_roll
            rep[alias.lower()] = std_roll
            rep[alias.upper()] = std_roll
            rep[alias.title()] = std_roll

    if std_email:
        for alias in ["Email", "Mail id", "Mail ID", "College mail id", "student_email"]:
            if alias not in rep:
                rep[alias] = std_email

    # 4. Format any date-like values in the replacements dictionary to dd-mm-yyyy with non-breaking hyphens
    for k, v in list(rep.items()):
        if isinstance(k, str) and ('date' in k.lower() or k.lower() == 'date'):
            rep[k] = format_date_to_dd_mm_yyyy(v)

    return rep


def replace_tokens_in_pptx_slide(slide, replacements: dict):
    """Replaces <<Placeholder>> tokens inside PowerPoint slides while preserving 100% of font family, font size, bold, italic, and colors."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        try:
            shape.text_frame.word_wrap = True
        except Exception:
            pass

        # Normalize fonts and disable underlines slide-wide
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.font:
                    r.font.underline = False
                    if r.font.name:
                        name_lower = r.font.name.lower()
                        if "times" in name_lower:
                            r.font.name = "Times New Roman"
                            if "bold" in name_lower:
                                r.font.bold = True
                        elif "playfair" in name_lower:
                            r.font.name = "Playfair Display"
                            if "bold" in name_lower:
                                r.font.bold = True

        for p in shape.text_frame.paragraphs:
            full_text = p.text
            if not full_text or '<<' not in full_text:
                continue

            # Find all <<token>> matches with character start/end indices in full_text
            matches = []
            for m in re.finditer(r"<<\s*([^>]+?)\s*>>", full_text):
                token_clean = m.group(1).strip()
                val = None
                for k in [m.group(1), token_clean, token_clean.lower(), token_clean.upper(), token_clean.title(), token_clean.replace('_', ' '), token_clean.replace(' ', '_')]:
                    if k in replacements:
                        val = replacements[k]
                        break
                if val is not None:
                    val_str = str(val)
                    # If the placeholder is followed immediately by an alphanumeric character (like "o" in "organized"),
                    # append a space to prevent breaking words and line wraps.
                    if m.end() < len(full_text) and full_text[m.end()].isalnum():
                        val_str += " "
                    matches.append((m.start(), m.end(), val_str))

            if not matches:
                continue

            # Map each character index in full_text to (run_index, char_index_within_run)
            char_map = []
            for r_idx, r in enumerate(p.runs):
                for c_idx in range(len(r.text)):
                    char_map.append((r_idx, c_idx))

            # Perform replacement from right to left so indices remain valid
            for start, end, val_str in reversed(matches):
                if start >= len(char_map) or end > len(char_map):
                    continue

                start_run_idx, start_char_offset = char_map[start]
                end_run_idx, end_char_offset = char_map[end - 1]

                if start_run_idx == end_run_idx:
                    # Token is inside a single run
                    r = p.runs[start_run_idx]
                    r.text = r.text[:start_char_offset] + val_str + r.text[end_char_offset + 1:]
                else:
                    # Token spans across multiple runs
                    r_first = p.runs[start_run_idx]
                    r_first.text = r_first.text[:start_char_offset] + val_str

                    for mid_idx in range(start_run_idx + 1, end_run_idx):
                        p.runs[mid_idx].text = ""

                    r_last = p.runs[end_run_idx]
                    r_last.text = r_last.text[end_char_offset + 1:]

def generate_single_native_pdf(pptx_template_path: str, replacements: dict, output_pdf_path: str) -> bool:
    """Modifies PPTX template with student replacements and converts directly to PDF using PowerPoint COM with CoInitialize."""
    target_pptx = pptx_template_path if (pptx_template_path and os.path.exists(pptx_template_path)) else get_active_pptx_template_path()
    
    if not target_pptx or not os.path.exists(target_pptx):
        print(f"[Generator] PPTX template not found at {target_pptx}")
        return False

    temp_dir = tempfile.mkdtemp()
    temp_pptx = os.path.join(temp_dir, f"temp_{uuid.uuid4().hex[:6]}.pptx")

    co_initialized = False
    try:
        prs = Presentation(target_pptx)
        slide = prs.slides[0]
        replace_tokens_in_pptx_slide(slide, replacements)
        prs.save(temp_pptx)

        if has_win32com:
            if pythoncom_lib:
                pythoncom_lib.CoInitialize()
                co_initialized = True

            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            pres = ppt_app.Presentations.Open(os.path.abspath(temp_pptx), WithWindow=False)
            pres.SaveAs(os.path.abspath(output_pdf_path), 32) # 32 = ppSaveAsPDF
            pres.Close()
            ppt_app.Quit()

            if co_initialized:
                pythoncom_lib.CoUninitialize()

            print(f"[Generator] Successfully rendered native PowerPoint PDF at {output_pdf_path}")
            return True
        else:
            print("[Generator] win32com unavailable. Unable to render native PDF.")
            return False
    except Exception as e:
        print(f"[Generator] Native PPTX to PDF error: {e}")
        if co_initialized and pythoncom_lib:
            try:
                pythoncom_lib.CoUninitialize()
            except Exception:
                pass
        return False

def pdf_to_base64_png(pdf_path: str) -> str:
    """Renders page 1 of PDF file to Base64 PNG data URL."""
    if not has_fitz or not os.path.exists(pdf_path):
        return ""
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        b64 = base64.b64encode(img_bytes).decode('utf-8')
        return f"data:image/png;base64,{b64}"
    except Exception as e:
        print(f"[Generator] PDF to PNG error: {e}")
        return ""

def parse_excel_dataframe(file_bytes: bytes, filename: str):
    """Parses Excel bytes into records dictionary."""
    if filename.endswith('.csv'):
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))

    df = df.fillna('')
    headers = list(df.columns)
    rows = df.to_dict(orient='records')
    return {
        "headers": headers,
        "rows": rows,
        "total_rows": len(rows)
    }
