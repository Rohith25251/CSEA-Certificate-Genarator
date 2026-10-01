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
    if os.name == 'nt':
        print(f"[Generator] win32com/pythoncom import notice: {err}")
    has_win32com = False

# Attempt PyMuPDF for PDF page rasterization
has_fitz = False
try:
    try:
        import pymupdf as fitz
    except ImportError:
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

is_vercel = os.environ.get('VERCEL') is not None or os.environ.get('AWS_LAMBDA_FUNCTION_NAME') is not None

if is_vercel:
    UPLOAD_DIR = os.path.join('/tmp', 'uploaded_templates')
else:
    UPLOAD_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), 'uploaded_templates'))
os.makedirs(UPLOAD_DIR, exist_ok=True)

SAVED_PPTX_PATH = os.path.join(UPLOAD_DIR, 'latest_template.pptx')

def get_active_pptx_template_path(event_id: str = None) -> str:
    """Returns path to active PPTX template, prioritizing local event_templates/{event_id}/ folder, then latest_template.pptx, then Supabase Storage."""
    # 1. First check local event_templates/{event_id}/
    if event_id:
        clean_eid = str(event_id).strip()
        candidates = [
            clean_eid,
            clean_eid.replace(' ', ''),
            clean_eid.replace(' ', '-'),
            clean_eid.replace('-', ' ').strip()
        ]
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "event_templates"))
        for cand in candidates:
            cand_dir = os.path.join(base_dir, cand)
            if os.path.exists(cand_dir):
                pptxs = [os.path.join(cand_dir, f) for f in os.listdir(cand_dir) if f.endswith('.pptx')]
                if pptxs:
                    pptxs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                    return pptxs[0]

        # 2. Check Supabase Storage under the specific event_id folder
        from supabase_client import download_latest_template_from_supabase
        for cand in candidates:
            sp_path = download_latest_template_from_supabase(cand)
            if sp_path and os.path.exists(sp_path):
                return sp_path

    # 3. Check latest uploaded template file from current session
    if os.path.exists(SAVED_PPTX_PATH):
        return SAVED_PPTX_PATH

    from supabase_client import download_latest_template_from_supabase
    sp_latest = download_latest_template_from_supabase()
    if sp_latest and os.path.exists(sp_latest):
        return sp_latest

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

def extract_row_title_from_dict(row: dict, fallback: str = "") -> str:
    """Extracts student's individual paper/project title or topic from Excel row."""
    if not row or not isinstance(row, dict):
        return fallback
    for pk in [
        'Title', 'title', 'TITLE', 'Title ', 'title ', 'Paper Title', 'paper_title', 
        'Project Title', 'project_title', 'Topic', 'topic', 'TOPIC', 'Topic ', 'topic '
    ]:
        if pk in row and row[pk] is not None and str(row[pk]).strip():
            return str(row[pk]).strip()
    for k, v in row.items():
        if not k or v is None:
            continue
        ck = str(k).strip().lower().replace('_', ' ')
        cv = str(v).strip()
        if not cv:
            continue
        if ('title' in ck or 'paper' in ck or 'project' in ck) and not ('event' in ck or 'date' in ck or 'id' in ck):
            return cv
    return fallback

def extract_row_event_name_from_dict(row: dict, fallback: str = "") -> str:
    """Extracts student's individual event name from arbitrary Excel row keys."""
    if not row or not isinstance(row, dict):
        return fallback
    
    # Priority exact keys for Event Name
    for pk in [
        'Event', 'event', 'EVENT', 'Event ', 'event ', 'Event Name', 'event_name', 'Event_Name', 'EVENT_NAME',
        'Event Title', 'event_title', 'Event_Title',
        'Workshop', 'workshop', 'Workshop Name', 'workshop_name',
        'Activity', 'activity', 'Competition', 'competition',
        'Track', 'track', 'Name of the Event', 'Name of the event'
    ]:
        if pk in row and row[pk] is not None and str(row[pk]).strip():
            return str(row[pk]).strip()
            
    # Fuzzy search across keys
    for k, v in row.items():
        if not k or v is None:
            continue
        ck = str(k).strip().lower().replace('_', ' ')
        cv = str(v).strip()
        if not cv:
            continue
        if (
            ('event' in ck or 'workshop' in ck or 'activity' in ck or 'competition' in ck or 'track' in ck) and
            not ('date' in ck or 'id' in ck or 'category' in ck or 'college' in ck or 'student' in ck or 'participant' in ck or 'mail' in ck)
        ):
            return cv
            
    return fallback

def build_dynamic_replacements(row: dict, extra: dict = None) -> dict:
    """
    Builds a universal dynamic replacement dictionary from ANY arbitrary Excel column header,
    handling trailing/leading spaces, casing, underscore/space variations, and common aliases.
    Prioritizes row-level Excel values over extra batch defaults.
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

    # 2. Extract standard fields (Name, Roll No, Email, Date, College, Event, Year, Section) from row
    std_name = None
    std_roll = None
    std_email = None
    std_event = extract_row_event_name_from_dict(row)
    std_year = None
    std_section = None
    std_college = None

    for k, v in row.items():
        if not k or v is None:
            continue
        ck = str(k).strip().lower().replace('_', ' ')
        cv = str(v).strip()
        if not cv:
            continue

        if ('name' in ck or 'student' in ck or 'participant' in ck) and not ('event' in ck or 'college' in ck or 'workshop' in ck) and not std_name:
            std_name = cv
        if ('roll' in ck or 'register' in ck or 'reg' in ck or 'id' in ck or 'number' in ck) and not ('phone' in ck or 'mobile' in ck or 'contact' in ck) and not std_roll:
            std_roll = cv
        if ('mail' in ck or 'email' in ck) and '@' in cv and not std_email:
            std_email = cv
        if ('year' in ck or 'academic year' in ck or 'batch' in ck) and not ('date' in ck) and not std_year:
            std_year = cv
        if ('section' in ck or 'sec' == ck) and not std_section:
            std_section = cv
        if ('college' in ck or 'institution' in ck) and not std_college:
            std_college = cv

    # Fallback to extra if row didn't have event/name/roll
    if not std_event and extra:
        std_event = extra.get('Event') or extra.get('event_name') or extra.get('Event Name') or extra.get('event')

    # 3. Register standard aliases for PowerPoint templates with varied token naming
    if std_name:
        for alias in ["Name", "Name ", "name", "Full Name", "Full Name ", "Student Name", "Participant Name", "PARTICIPANT_NAME"]:
            rep[alias] = std_name
            rep[alias.lower()] = std_name
            rep[alias.upper()] = std_name

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
        for alias in ["Email", "Mail id", "Mail ID", "College mail id", "student_email", "Email Address"]:
            rep[alias] = std_email

    std_title = extract_row_title_from_dict(row)
    if std_title:
        for alias in ["Title", "Title ", "title", "TITLE", "Topic", "topic", "TOPIC", "Paper Title", "paper_title", "Project Title"]:
            rep[alias] = std_title
            rep[alias.lower()] = std_title
            rep[alias.upper()] = std_title

    if std_event:
        for alias in [
            "event_name", "Event Name", "Event Name ", "event name", "Event_Name", "EVENT_NAME",
            "event", "Event", "Event ", "EVENT", "Event Title", "event_title", "Event_Title",
            "Workshop", "Workshop Name", "workshop_name",
            "Name of the Event", "Name of the event"
        ]:
            rep[alias] = std_event
            rep[alias.lower()] = std_event
            rep[alias.upper()] = std_event
            rep[alias.title()] = std_event
            rep[alias.title()] = std_event

    if std_year:
        for alias in [
            "Year", "Year ", "year", "YEAR", "Year of Study", "Year of Study ", 
            "year_of_study", "Year_Of_Study", "YEAR_OF_STUDY", "Year of study",
            "Academic Year", "academic_year", "Batch", "batch"
        ]:
            rep[alias] = std_year
            rep[alias.lower()] = std_year
            rep[alias.upper()] = std_year
            rep[alias.title()] = std_year

    if std_section:
        for alias in ["Section", "section", "SECTION", "Sec", "sec"]:
            rep[alias] = std_section

    if std_college:
        for alias in ["College", "College Name", "college_name", "COLLEGE", "college"]:
            rep[alias] = std_college

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

                    # If any run within the token was underlined, ensure r_first keeps the underline
                    was_underlined = any(
                        p.runs[idx].font.underline for idx in range(start_run_idx, end_run_idx + 1) 
                        if p.runs[idx].font and p.runs[idx].font.underline
                    )
                    if was_underlined and r_first.font:
                        r_first.font.underline = True

                    for mid_idx in range(start_run_idx + 1, end_run_idx):
                        p.runs[mid_idx].text = ""

                    r_last = p.runs[end_run_idx]
                    r_last.text = r_last.text[end_char_offset + 1:]

def replace_html_tokens(html_content: str, replacements: dict) -> str:
    """Replaces <<placeholder>> and {{placeholder}} in HTML content with replacement values (case-insensitive)."""
    rendered = html_content
    for k, v in replacements.items():
        val = str(v) if v is not None else ""
        escaped_k = re.escape(str(k).strip())
        pattern = r"(<<\s*" + escaped_k + r"\s*>>|\{\{\s*" + escaped_k + r"\s*\}\})"
        rendered = re.sub(pattern, val, rendered, flags=re.IGNORECASE)
    return rendered

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
            print("[Generator] win32com unavailable. Falling back to high-fidelity ReportLab DrawingML PDF rendering.")
            try:
                import io
                from pptx.enum.shapes import MSO_SHAPE_TYPE
                from pptx.enum.text import PP_ALIGN
                from reportlab.pdfgen import canvas
                from reportlab.lib.pagesizes import A4, landscape
                from reportlab.platypus import Paragraph
                from reportlab.lib.styles import ParagraphStyle
                from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
                from reportlab.lib.utils import ImageReader
                from reportlab.lib.colors import HexColor
                from svglib.svglib import svg2rlg
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.ttfonts import TTFont
                from PIL import Image

                # Register bundled Fonts
                fonts_dir = os.path.join(os.path.dirname(__file__), "fonts")
                reg_fonts = set()
                font_map = {
                    "LeagueGothic": "LeagueGothic-Regular.ttf",
                    "Poppins": "Poppins-Regular.ttf",
                    "Poppins-Bold": "Poppins-Bold.ttf",
                    "Poppins-Medium": "Poppins-Medium.ttf",
                    "Poppins-SemiBold": "Poppins-SemiBold.ttf",
                    "PaytoneOne": "PaytoneOne-Regular.ttf",
                    "Righteous": "Righteous-Regular.ttf",
                    "Montserrat-Bold": "Montserrat-Bold.ttf",
                    "Anton": "Anton-Regular.ttf",
                    "BebasNeue": "BebasNeue-Regular.ttf",
                    "Bungee": "Bungee-Regular.ttf",
                    "Chonburi": "Chonburi-Regular.ttf",
                    "Shrikhand": "Shrikhand-Regular.ttf",
                    "TitanOne": "TitanOne-Regular.ttf"
                }
                for f_name, f_file in font_map.items():
                    fp = os.path.join(fonts_dir, f_file)
                    if os.path.exists(fp):
                        try:
                            pdfmetrics.registerFont(TTFont(f_name, fp))
                            reg_fonts.add(f_name)
                        except Exception:
                            pass

                # 1. Clean replacements to remove any Unicode non-breaking hyphens
                cleaned_replacements = {}
                for k, v in replacements.items():
                    if isinstance(v, str):
                        cleaned_replacements[k] = v.replace('\u2011', '-')
                    else:
                        cleaned_replacements[k] = v

                # Use the modified temporary PPTX
                prs_temp = Presentation(temp_pptx)
                slide_temp = prs_temp.slides[0]

                # A4 Page dimensions in points
                A4_w, A4_h = landscape(A4) # 841.89 x 595.27

                os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
                pdf_canvas = canvas.Canvas(output_pdf_path, pagesize=(A4_w, A4_h))

                # Set PDF document metadata so browser tab titles display the student's name and event
                student_title_name = str(replacements.get('Name') or replacements.get('name') or replacements.get('Student Name') or 'Participant').strip()
                event_title_name = str(replacements.get('event_name') or replacements.get('Event Name') or replacements.get('Event') or 'Certificate').strip()
                pdf_canvas.setTitle(f"Certificate - {student_title_name} - {event_title_name}")
                pdf_canvas.setAuthor("CSEA - Kongu Engineering College")
                pdf_canvas.setSubject(f"Certificate of Participation - {event_title_name}")

                slide_w = prs_temp.slide_width
                slide_h = prs_temp.slide_height
                scale_x = A4_w / slide_w
                scale_y = A4_h / slide_h

                # Slide Background
                bg_color = None
                try:
                    bg_el = slide_temp._element.xpath('.//p:bg//a:srgbClr')
                    if bg_el:
                        bg_color = f"#{bg_el[0].get('val')}"
                except Exception:
                    pass
                if bg_color:
                    pdf_canvas.setFillColor(HexColor(bg_color))
                    pdf_canvas.rect(0, 0, A4_w, A4_h, fill=1, stroke=0)

                def extract_image_bytes(shape, slide_part):
                    try:
                        if hasattr(shape, 'image') and shape.image:
                            return shape.image.blob, getattr(shape.image, 'filename', '')
                    except Exception:
                        pass
                    try:
                        rIds = shape._element.xpath('.//@r:embed')
                        if rIds:
                            part = slide_part.related_part(rIds[0])
                            ext = getattr(part, 'partname', '')
                            return part.blob, str(ext)
                    except Exception:
                        pass
                    return None, ''

                def get_shape_fill_color(shape):
                    try:
                        fills = shape._element.xpath('./*[local-name()="spPr"]/*[local-name()="solidFill"]/*[local-name()="srgbClr"]')
                        if fills:
                            return f"#{fills[0].get('val')}"
                    except Exception:
                        pass
                    return None

                def get_shape_line_info(shape):
                    try:
                        ln = shape._element.xpath('./*[local-name()="spPr"]/*[local-name()="ln"]')
                        if ln:
                            w_val = ln[0].get('w')
                            lw = (int(w_val) / 12700.0) if w_val else 1.0
                            clrs = ln[0].xpath('.//*[local-name()="srgbClr"]')
                            color_hex = f"#{clrs[0].get('val')}" if clrs else None
                            has_head_oval = bool(ln[0].xpath('.//*[local-name()="headEnd"][@type="oval"]'))
                            has_tail_oval = bool(ln[0].xpath('.//*[local-name()="tailEnd"][@type="oval"]'))
                            return color_hex, lw, has_head_oval, has_tail_oval
                    except Exception:
                        pass
                    return None, 1.0, False, False

                def select_font(raw_font_name, is_bold, is_italic):
                    fn = (raw_font_name or "").lower()
                    if "league gothic" in fn or "gothic" in fn:
                        return "LeagueGothic" if "LeagueGothic" in reg_fonts else "Helvetica-Bold"
                    if "motter" in fn or "corpus" in fn or "paytone" in fn or "renaissance" in fn:
                        return "PaytoneOne" if "PaytoneOne" in reg_fonts else "Helvetica-Bold"
                    if "righteous" in fn:
                        return "Righteous" if "Righteous" in reg_fonts else "Helvetica-Bold"
                    if "tt hoves" in fn or "hoves" in fn:
                        return "Poppins-Bold" if "Poppins-Bold" in reg_fonts else "Helvetica-Bold"
                    if "montserrat" in fn:
                        return "Poppins-Bold" if "Poppins-Bold" in reg_fonts else "Helvetica-Bold"
                    if "poppins" in fn:
                        if is_bold:
                            return "Poppins-Bold" if "Poppins-Bold" in reg_fonts else "Helvetica-Bold"
                        return "Poppins" if "Poppins" in reg_fonts else "Helvetica"
                    if "times" in fn or "playfair" in fn or "serif" in fn:
                        if is_bold and is_italic:
                            return "Times-BoldItalic"
                        if is_bold:
                            return "Times-Bold"
                        if is_italic:
                            return "Times-Italic"
                        return "Times-Roman"
                    
                    if is_bold and is_italic:
                        return "Helvetica-BoldOblique"
                    if is_bold:
                        return "Helvetica-Bold"
                    if is_italic:
                        return "Helvetica-Oblique"
                    return "Helvetica"

                def render_element(shape, abs_left, abs_top, abs_w, abs_h):
                    # If it is a group, recurse its children with exact DrawingML group coordinate mapping
                    if hasattr(shape, 'shapes'):
                        try:
                            xfrm = shape._element.grpSpPr.xfrm
                            chOff_x = int(xfrm.chOff.x) if hasattr(xfrm, 'chOff') and hasattr(xfrm.chOff, 'x') else 0
                            chOff_y = int(xfrm.chOff.y) if hasattr(xfrm, 'chOff') and hasattr(xfrm.chOff, 'y') else 0
                            chExt_x = int(xfrm.chExt.cx) if hasattr(xfrm, 'chExt') and hasattr(xfrm.chExt, 'cx') else shape.width
                            chExt_y = int(xfrm.chExt.cy) if hasattr(xfrm, 'chExt') and hasattr(xfrm.chExt, 'cy') else shape.height

                            grp_scale_x = abs_w / chExt_x if chExt_x else 1.0
                            grp_scale_y = abs_h / chExt_y if chExt_y else 1.0

                            for sub in shape.shapes:
                                sub_left = abs_left + (sub.left - chOff_x) * grp_scale_x
                                sub_top = abs_top + (sub.top - chOff_y) * grp_scale_y
                                sub_w = sub.width * grp_scale_x
                                sub_h = sub.height * grp_scale_y
                                render_element(sub, sub_left, sub_top, sub_w, sub_h)
                        except Exception as grp_err:
                            print(f"[Generator] Group render error: {grp_err}")
                        return

                    x = abs_left * scale_x
                    w = abs_w * scale_x
                    h = abs_h * scale_y
                    top_y = A4_h - abs_top * scale_y
                    bot_y = top_y - h

                    # 1. Solid fill on shapes (e.g. Ivory background card)
                    fill_hex = get_shape_fill_color(shape)
                    img_bytes, fname = extract_image_bytes(shape, slide_temp._part)
                    
                    if fill_hex and not img_bytes:
                        pdf_canvas.setFillColor(HexColor(fill_hex))
                        pdf_canvas.rect(x, bot_y, w, h, fill=1, stroke=0)

                    # 2. Check for image/picture fill / SVGs
                    if img_bytes:
                        try:
                            xfrm_el = shape._element.xpath('.//*[local-name()="xfrm"]')
                            flip_h = False
                            flip_v = False
                            if xfrm_el:
                                flip_h = xfrm_el[0].get('flipH') in ['1', 'true']
                                flip_v = xfrm_el[0].get('flipV') in ['1', 'true']

                            is_svg = fname.lower().endswith('.svg') or img_bytes.startswith(b'<svg') or b'<svg' in img_bytes[:200]
                            if is_svg:
                                drawing = svg2rlg(io.BytesIO(img_bytes))
                                sx = w / drawing.width
                                sy = h / drawing.height
                                pdf_canvas.saveState()
                                if flip_h or flip_v:
                                    pdf_canvas.translate(x + (w if flip_h else 0), bot_y + (h if flip_v else 0))
                                    pdf_canvas.scale(-sx if flip_h else sx, -sy if flip_v else sy)
                                    drawing.drawOn(pdf_canvas, 0, 0)
                                else:
                                    drawing.scale(sx, sy)
                                    drawing.drawOn(pdf_canvas, x, bot_y)
                                pdf_canvas.restoreState()
                            else:
                                im = Image.open(io.BytesIO(img_bytes))
                                if flip_h:
                                    im = im.transpose(Image.FLIP_LEFT_RIGHT)
                                if flip_v:
                                    im = im.transpose(Image.FLIP_TOP_BOTTOM)
                                
                                buf = io.BytesIO()
                                im.save(buf, format="PNG")
                                buf.seek(0)
                                img_reader = ImageReader(buf)
                                pdf_canvas.drawImage(img_reader, x, bot_y, w, h, mask='auto')
                        except Exception as e:
                            print(f"[Generator] Image render error on {getattr(shape, 'name', '')}: {e}")

                    # 3. Check for line dividers and end circle dots
                    if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and abs_h == 0:
                        line_color, lw, has_head_oval, has_tail_oval = get_shape_line_info(shape)
                        if line_color:
                            pdf_canvas.setStrokeColor(HexColor(line_color))
                            pdf_canvas.setFillColor(HexColor(line_color))
                            pdf_canvas.setLineWidth(lw)
                            pdf_canvas.line(x, top_y, x + w, top_y)
                            dot_r = max(2.5, lw * 3.0)
                            if has_head_oval:
                                pdf_canvas.circle(x, top_y, dot_r, fill=1, stroke=0)
                            if has_tail_oval:
                                pdf_canvas.circle(x + w, top_y, dot_r, fill=1, stroke=0)

                    # 4. Check for text frame
                    if hasattr(shape, 'text_frame') and shape.has_text_frame and shape.text_frame.text.strip():
                        curr_y = top_y
                        for p in shape.text_frame.paragraphs:
                            if not p.text.strip():
                                continue
                            p_text = ''

                            # Check for paragraph line spacing
                            spc_pts = p._p.xpath('.//a:lnSpc/a:spcPts/@val')
                            spc_pct = p._p.xpath('.//a:lnSpc/a:spcPct/@val')
                            paragraph_leading = None
                            if spc_pts:
                                paragraph_leading = int(spc_pts[0]) / 100.0
                            elif spc_pct:
                                paragraph_leading = None
                            elif p.line_spacing:
                                if isinstance(p.line_spacing, (int, float)) and p.line_spacing > 50:
                                    paragraph_leading = p.line_spacing / 12700.0

                            # Paragraph default font size
                            def_sz_val = p._p.xpath('./a:pPr/a:defRPr/@sz')
                            p_def_sz = (int(def_sz_val[0]) / 100.0) if def_sz_val else 18.0

                            current_p_max_sz = p_def_sz
                            has_p_underline = False

                            for r in p.runs:
                                t = r.text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                                t = t.replace('\u2011', '-')
                                if not t:
                                    continue
                                
                                is_bold = bool(r.font.bold)
                                is_italic = bool(r.font.italic)
                                rl_font = select_font(r.font.name, is_bold, is_italic)

                                # Determine font size
                                r_sz_xml = r._r.xpath('./a:rPr/@sz')
                                spc_xml = r._r.xpath('./a:rPr/@spc')
                                if r_sz_xml:
                                    sz = int(r_sz_xml[0]) / 100.0
                                elif r.font.size and hasattr(r.font.size, 'pt'):
                                    sz = r.font.size.pt
                                else:
                                    sz = p_def_sz

                                # If PPTX had negative character spacing (e.g. spc="-77"), slightly scale font size to 17.0pt for exact fit
                                if spc_xml and int(spc_xml[0]) < 0 and sz == 18.0:
                                    sz = 17.0

                                if sz > current_p_max_sz:
                                    current_p_max_sz = sz

                                c_hex = '#000000'
                                try:
                                    if r.font.color and r.font.color.type == 1:
                                        rgb = r.font.color.rgb
                                        c_hex = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
                                    else:
                                        r_clrs = r._r.xpath('.//*[local-name()="srgbClr"]')
                                        if r_clrs:
                                            c_hex = f"#{r_clrs[0].get('val')}"
                                except Exception:
                                    pass

                                # Check for underline in r.font.underline or OpenXML u attribute
                                is_underlined = bool(r.font.underline)
                                if not is_underlined:
                                    u_attr = r._r.xpath('.//@u')
                                    if u_attr and u_attr[0] not in ['none', '0']:
                                        is_underlined = True

                                style_s = f'<font name="{rl_font}" size="{sz:.1f}" color="{c_hex}">'
                                style_e = '</font>'
                                if is_underlined:
                                    has_p_underline = True
                                    style_s += '<u>'
                                    style_e = '</u>' + style_e
                                p_text += f'{style_s}{t}{style_e}'

                            align = TA_LEFT
                            if p.alignment == PP_ALIGN.CENTER:
                                align = TA_CENTER
                            elif p.alignment == PP_ALIGN.RIGHT:
                                align = TA_RIGHT
                            elif p.alignment == PP_ALIGN.JUSTIFY:
                                align = TA_JUSTIFY

                            effective_leading = paragraph_leading if paragraph_leading else (current_p_max_sz * 1.35)

                            p_style = ParagraphStyle(
                                name=f's_{uuid.uuid4().hex[:6]}',
                                alignment=align,
                                leading=effective_leading
                            )
                            para = Paragraph(p_text, p_style)

                            pw, ph = para.wrap(w, A4_h)

                            pdf_canvas.saveState()
                            if has_p_underline:
                                pdf_canvas.setDash(2.5, 2.0)
                                pdf_canvas.setLineWidth(0.75)

                            para.drawOn(pdf_canvas, x, curr_y - ph)
                            pdf_canvas.restoreState()

                            curr_y -= ph

                for s in slide_temp.shapes:
                    render_element(s, s.left, s.top, s.width, s.height)

                pdf_canvas.showPage()
                pdf_canvas.save()
                print(f"[Generator] Successfully rendered high-fidelity fallback PDF using ReportLab at {output_pdf_path}")
                return True
            except Exception as fallback_err:
                print(f"[Generator] Fallback ReportLab rendering failed: {fallback_err}")
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
