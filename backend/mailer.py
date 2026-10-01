import os
import re
import smtplib
import base64
from typing import Optional
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.message import EmailMessage
from supabase_client import supabase

def get_smtp_config():
    """Loads SMTP environment variables supporting both SMTP_SERVER/SMTP_USERNAME and SMTP_HOST/SMTP_USER."""
    host = os.getenv("SMTP_SERVER") or os.getenv("SMTP_HOST") or ""
    port = int(os.getenv("SMTP_PORT") or "587")
    user = os.getenv("SMTP_USERNAME") or os.getenv("SMTP_USER") or ""
    password = os.getenv("SMTP_PASSWORD") or os.getenv("SMTP_PASS") or ""
    
    from_name = os.getenv("SMTP_FROM_NAME", "CSEA Association")
    from_email = os.getenv("SMTP_FROM_EMAIL") or os.getenv("SMTP_FROM") or user
    
    if "<" in from_email:
        from_header = from_email
    else:
        from_header = f"{from_name} <{from_email}>"

    return {
        "host": host,
        "port": port,
        "user": user,
        "pass": password,
        "password": password,
        "from": from_header,
        "hero_img_url": os.getenv("EMAIL_HERO_IMAGE_URL", "http://localhost:3000/hero.png"),
        "logo_img_url": os.getenv("EMAIL_LOGO_IMAGE_URL", "http://localhost:3000/csea_logo.png")
    }

def generate_certificate_email_html(
    student_name: str,
    event_name: str = "Workshop",
    event_date: str = "2026-07-25",
    custom_hero_url: Optional[str] = None,
    custom_logo_url: Optional[str] = None,
    custom_template_html: Optional[str] = None
) -> str:
    """Generates the full styled CSEA HTML email body with dynamic recipient, event name, and date details."""
    # Reformat date from YYYY-MM-DD to DD-Mon-YYYY (e.g. 2026-08-01 -> 01-Aug-2026)
    try:
        from datetime import datetime
        parsed_date = datetime.strptime(str(event_date).strip(), "%Y-%m-%d")
        event_date = parsed_date.strftime("%d-%b-%Y")
    except Exception:
        pass  # Keep original string if parsing fails

    supabase_url = os.getenv('NEXT_PUBLIC_SUPABASE_URL') or "https://bqvnuvfmddtyvpxuceol.supabase.co"
    default_logo_src = f"{supabase_url}/storage/v1/object/public/templates/assets/email_logo.png"
    default_hero_src = f"{supabase_url}/storage/v1/object/public/templates/assets/email_hero.png"

    logo_src = custom_logo_url
    hero_src = custom_hero_url

    # Check Supabase app_settings for custom logo/hero if not passed explicitly
    if supabase and (not logo_src or not hero_src):
        try:
            res_settings = supabase.table('app_settings').select('*').in_('key', ['use_custom_logo', 'use_custom_hero']).execute()
            if res_settings.data:
                for row in res_settings.data:
                    k = row.get('key')
                    v = row.get('value')
                    if k == 'use_custom_logo' and not logo_src:
                        if isinstance(v, dict) and v.get('enabled') and v.get('url'):
                            logo_src = v.get('url')
                        elif isinstance(v, str) and v.startswith('http'):
                            logo_src = v
                    elif k == 'use_custom_hero' and not hero_src:
                        if isinstance(v, dict) and v.get('enabled') and v.get('url'):
                            hero_src = v.get('url')
                        elif isinstance(v, str) and v.startswith('http'):
                            hero_src = v
        except Exception as e:
            print(f"[Supabase] Settings fetch warning in mailer: {e}")

    logo_src = logo_src or default_logo_src
    hero_src = hero_src or default_hero_src

    # 1. Check if custom template is passed or saved in Supabase app_settings
    tpl = custom_template_html
    if not tpl and supabase:
        try:
            res = supabase.table('app_settings').select('value').eq('key', 'custom_email_template').execute()
            if res.data and len(res.data) > 0:
                val = res.data[0].get('value')
                if isinstance(val, dict) and val.get('enabled') and val.get('html'):
                    tpl = val.get('html')
                elif isinstance(val, str) and len(val.strip()) > 50:
                    tpl = val
        except Exception:
            pass

    if tpl and len(tpl.strip()) > 50:
        rendered = tpl
        replacements = {
            "student_name": student_name,
            "name": student_name,
            "Name": student_name,
            "event_name": event_name,
            "Event Name": event_name,
            "Event_Name": event_name,
            "event": event_name,
            "Event": event_name,
            "EVENT": event_name,
            "Workshop": event_name,
            "Topic": event_name,
            "event_date": event_date,
            "Date": event_date,
            "Event Date": event_date,
            "hero_src": hero_src,
            "logo_src": logo_src
        }
        for k, v in replacements.items():
            rendered = rendered.replace(f"{{{k}}}", str(v))
            rendered = rendered.replace(f"{{{{{k}}}}}", str(v))
            rendered = rendered.replace(f"<<{k}>>", str(v))

        # Replace default URL instances if custom hero/logo is set
        if hero_src and hero_src != default_hero_src:
            rendered = rendered.replace(default_hero_src, hero_src)
            rendered = rendered.replace("/hero.png", hero_src)
            rendered = rendered.replace("http://localhost:3000/hero.png", hero_src)
            rendered = rendered.replace("http://localhost:3000/email_hero.png", hero_src)
            rendered = re.sub(r'(<img[^>]*?alt=["\'](?:CSEA Header Hero|Hero Banner|Hero Image)["\'][^>]*?src=["\'])[^"\']+["\']', rf'\g<1>{hero_src}"', rendered, flags=re.IGNORECASE)
            rendered = re.sub(r'(<img[^>]*?src=["\'])[^"\']+["\']([^>]*?alt=["\'](?:CSEA Header Hero|Hero Banner|Hero Image)["\'])', rf'\g<1>{hero_src}"\g<2>', rendered, flags=re.IGNORECASE)

        if logo_src and logo_src != default_logo_src:
            rendered = rendered.replace(default_logo_src, logo_src)
            rendered = rendered.replace("/csea_logo.png", logo_src)
            rendered = rendered.replace("http://localhost:3000/csea_logo.png", logo_src)
            rendered = rendered.replace("http://localhost:3000/email_logo.png", logo_src)
            rendered = re.sub(r'(<img[^>]*?alt=["\'](?:CSEA Logo|Logo)["\'][^>]*?src=["\'])[^"\']+["\']', rf'\g<1>{logo_src}"', rendered, flags=re.IGNORECASE)
            rendered = re.sub(r'(<img[^>]*?src=["\'])[^"\']+["\']([^>]*?alt=["\'](?:CSEA Logo|Logo)["\'])', rf'\g<1>{logo_src}"\g<2>', rendered, flags=re.IGNORECASE)

        return rendered

    return f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <style>
    @media screen and (max-width: 600px) {{
      .header-container {{
        flex-direction: column !important;
        text-align: center !important;
      }}
      .logo-wrapper {{
        margin-right: 0 !important;
        margin-bottom: 15px !important;
      }}
      .content-wrapper {{
        padding: 24px !important;
      }}
      .footer-cols {{
        flex-direction: column !important;
      }}
      .footer-col {{
        width: 100% !important;
        margin-bottom: 20px !important;
      }}
    }}
  </style>
</head>
<body style="font-family: 'Segoe UI', Arial, sans-serif; color: #333333; background-color: #FAF3E7; margin: 0; padding: 0;">

  <!-- MAIN CONTAINER -->
  <div style="max-width: 600px; margin: 40px auto; background-color: #FFFFFF; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08); border: 1px solid #FDE6DA;">
    
    <!-- TOP BROWSER-STYLE ACCENT BAR -->
    <div style="background-color: #1A1F2B; height: 10px; width: 100%;"></div>

    <!-- HERO IMAGE -->
    <div style="width: 100%; border-bottom: 2px solid #FAF3E7;">
      <img src="{hero_src}" alt="CSEA Header Hero" style="width: 100%; height: auto; display: block;" />
    </div>

    <!-- CONTENT WRAPPER -->
    <div class="content-wrapper" style="padding: 40px;">
      
      <!-- HEADER BRANDING & LOGO -->
      <div class="header-container" style="display: flex; align-items: center; margin-bottom: 30px;">
        <div class="logo-wrapper" style="margin-right: 24px; flex-shrink: 0;">
          <!-- ENLARGED CSEA LOGO -->
          <img src="{logo_src}" alt="CSEA Logo" style="width: 120px; max-width: 100%; height: auto; display: block;" />
        </div>
        <div>
          <h3 style="margin: 0; color: #1A1F2B; font-size: 18px; font-weight: 800; letter-spacing: -0.3px; line-height: 1.3;">
            COMPUTER SCIENCE ENGINEERING ASSOCIATION
          </h3>
          <p style="margin: 6px 0 0 0; font-size: 12px; color: #F15A24; font-weight: 700; text-transform: uppercase; letter-spacing: 1px;">
            Kongu Engineering College
          </p>
          <!-- BLUE ACCENT LINE -->
          <div style="width: 40px; height: 3px; background-color: #3B6FE0; margin-top: 8px; border-radius: 2px;"></div>
        </div>
      </div>

      <hr style="border: none; border-top: 1px dashed #FDE6DA; margin: 0 0 30px 0;" />

      <!-- MAIN BODY -->
      <h2 style="color: #1A1F2B; font-size: 22px; font-weight: 700; margin-top: 0;">
        Greetings, {student_name}! 🌟
      </h2>

      <p style="font-size: 15px; color: #333333; line-height: 1.7; margin-bottom: 20px;">
        Thank you for being an active participant in our technical endeavors! On behalf of the Computer Science Engineering Association (CSEA), we commend your dedication and passion for continuous learning.
      </p>

      <p style="font-size: 15px; color: #333333; line-height: 1.7; margin-bottom: 25px;">
        We are delighted to present your official participation details below:
      </p>
      
      <!-- EVENT DETAILS BOX -->
      <div style="background-color: #FCF3D9; padding: 20px; border-radius: 12px; margin: 25px 0; border-left: 4px solid #F15A24;">
        <p style="margin: 0; font-size: 14px; color: #1A1F2B;">
          <strong>Event Name:</strong> 
          <span style="font-weight: 700; color: #F15A24; margin-left: 5px;">
            {event_name}
          </span>
        </p>
        <p style="margin: 10px 0 0 0; font-size: 14px; color: #1A1F2B;">
          <strong>Event Date:</strong> 
          <span style="font-weight: 600; color: #333333; margin-left: 5px;">
            {event_date}
          </span>
        </p>
      </div>

      <p style="font-size: 14px; color: #6B7280; line-height: 1.6; margin-top: 25px;">
        We hope this experience was insightful and inspiring. We look forward to seeing your active participation in our upcoming workshops, hackathons, and technical symposiums!
      </p>

    </div> <!-- END CONTENT WRAPPER -->

    <!-- FOOTER - DARK BACKGROUND -->
    <div style="background-color: #111827; padding: 30px; border-top: 1px solid #1F2937;">
      <div class="footer-cols" style="display: flex; justify-content: space-between; gap: 20px;">
        
        <!-- FOOTER COL 1 -->
        <div class="footer-col" style="width: 50%;">
          <h4 style="color: #FFFFFF; font-size: 14px; font-weight: 700; margin: 0 0 10px 0;">
            About CSEA
          </h4>
          <p style="font-size: 12px; color: #9CA3AF; line-height: 1.6; margin: 0;">
            Computer Science and Engineering Association (CSEA) is the leading technical forum of Kongu Engineering College, nurturing student potential since inception.
          </p>
        </div>

        <!-- FOOTER COL 2 -->
        <div class="footer-col" style="width: 50%;">
          <h4 style="color: #F15A24; font-size: 14px; font-weight: 700; margin: 0 0 10px 0;">
            Contact & Location
          </h4>
          <p style="font-size: 12px; color: #9CA3AF; line-height: 1.6; margin: 0 0 8px 0;">
            📍 Department of CSE, Kongu Engineering College, Perundurai, Erode - 638060, Tamil Nadu, India.
          </p>
          <p style="font-size: 12px; color: #9CA3AF; line-height: 1.6; margin: 0 0 8px 0;">
            📞 +91 863 742 2716<br />
            📞 +91 861 099 2902
          </p>
          <p style="font-size: 12px; color: #9CA3AF; line-height: 1.6; margin: 0;">
            ✉️ <a href="mailto:csea@kongu.edu" style="color: #F15A24; text-decoration: none; font-weight: 600;">csea@kongu.edu</a>
          </p>
        </div>

      </div>

      <hr style="border: none; border-top: 1px solid #1F2937; margin: 25px 0 15px 0;" />
      
      <div style="text-align: center;">
        <span style="font-size: 11px; font-weight: 800; color: #FBBF24; letter-spacing: 1.5px;">WE CAN ∞ WE WILL</span>
        <p style="font-size: 11px; color: #9CA3AF; margin: 6px 0 0 0; line-height: 1.5;">
          © 2026 CSEA - Kongu Engineering College. All Rights Reserved.
        </p>
      </div>

    </div> <!-- END FOOTER -->

  </div> <!-- END MAIN CONTAINER -->

</body>
</html>
"""

def send_certificate_email(
    recipient_email: str,
    student_name: str,
    cert_code: str,
    pdf_path: str,
    event_name: str = "Workshop",
    event_date: str = "2026-07-25",
    custom_hero_url: Optional[str] = None,
    custom_logo_url: Optional[str] = None,
    cert_id: Optional[str] = None
) -> dict:
    """
    Sends certificate PDF via SMTP using the customized CSEA HTML email template.
    Returns the exact rendered HTML code alongside send status.
    """
    if not recipient_email or "@" not in recipient_email:
        return {"success": False, "error": "Invalid recipient email address"}

    cfg = get_smtp_config()
    
    msg = EmailMessage()
    msg['From'] = cfg["from"]
    msg['To'] = recipient_email
    msg['Subject'] = f"Certificate of Participation - {event_name} | CSEA"

    html_content = generate_certificate_email_html(
        student_name=student_name,
        event_name=event_name,
        event_date=event_date,
        custom_hero_url=custom_hero_url,
        custom_logo_url=custom_logo_url
    )

    # Set HTML content directly as the email body
    msg.set_content(html_content, subtype='html')

    # Attach PDF file if exists
    if pdf_path and os.path.exists(pdf_path):
        try:
            with open(pdf_path, 'rb') as f:
                pdf_data = f.read()
            attachment_filename = os.path.basename(pdf_path)
            msg.add_attachment(
                pdf_data,
                maintype='application',
                subtype='pdf',
                filename=attachment_filename
            )
        except Exception as e:
            print(f"[Mailer] Failed to add PDF attachment: {e}")
    else:
        print(f"[Mailer Notice] PDF attachment not found at {pdf_path}, sending email without attachment.")

    # Send Email via SMTP
    try:
        if cfg["pass"]:
            server = smtplib.SMTP(cfg["host"], cfg["port"])
            server.starttls()
            server.login(cfg["user"], cfg["pass"])
            server.send_message(msg)
            server.quit()
            print(f"[SMTP] Successfully dispatched custom CSEA email to {recipient_email}")
        else:
            print(f"[SMTP Simulation] Email prepared for {recipient_email} (Set SMTP_PASS in .env.local to deliver live)")

        # Update Supabase email_status to 'sent'
        if supabase:
            try:
                if cert_id:
                    supabase.table('certificates').update({"email_status": "sent"}).eq('id', cert_id).execute()
                    print(f"[Supabase] Updated email_status to 'sent' for certificate ID: {cert_id}")
                else:
                    # Fallback lookup: find student ID by email
                    s_res = supabase.table('students').select('id').eq('email', recipient_email).execute()
                    if s_res.data:
                        s_id = s_res.data[0]['id']
                        supabase.table('certificates').update({"email_status": "sent"}).eq('student_id', s_id).execute()
                        print(f"[Supabase] Updated email_status to 'sent' for student ID: {s_id}")
            except Exception as db_err:
                print(f"[Supabase] DB update notice: {db_err}")

        # Clean up local PDF file to save space and protect privacy
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
                print(f"[Mailer] Cleaned up temporary PDF: {pdf_path}")
            except Exception as clean_err:
                print(f"[Mailer] Failed to delete temporary PDF {pdf_path}: {clean_err}")

        return {
            "success": True, 
            "email": recipient_email, 
            "simulated": not bool(cfg["pass"]),
            "html_content": html_content
        }
    except Exception as e:
        print(f"[SMTP Error] Failed to send email to {recipient_email}: {e}")
        return {
            "success": False, 
            "email": recipient_email, 
            "error": str(e),
            "html_content": html_content
        }
    except Exception as e:
        print(f"[SMTP Error] Failed to send email to {recipient_email}: {e}")
        return {"success": False, "email": recipient_email, "error": str(e)}

def send_html_email(
    recipient_email: str,
    subject: str,
    html_content: str
) -> dict:
    """
    Sends pure HTML email via SMTP using the project SMTP configuration.
    """
    if not recipient_email or "@" not in recipient_email:
        return {"success": False, "error": "Invalid recipient email address"}

    cfg = get_smtp_config()
    msg = EmailMessage()
    msg['From'] = cfg["from"]
    msg['To'] = recipient_email
    msg['Subject'] = subject

    msg.set_content(html_content, subtype='html')

    try:
        if cfg["pass"]:
            server = smtplib.SMTP(cfg["host"], cfg["port"])
            server.starttls()
            server.login(cfg["user"], cfg["pass"])
            server.send_message(msg)
            server.quit()
            print(f"[SMTP] Successfully dispatched HTML invitation to {recipient_email}")
        else:
            print(f"[SMTP Simulation] HTML email prepared for {recipient_email} (Set SMTP_PASS in .env.local to deliver live)")

        return {"success": True, "email": recipient_email, "simulated": not bool(cfg["pass"])}
    except Exception as e:
        print(f"[SMTP Error] Failed to send HTML email to {recipient_email}: {e}")
        return {"success": False, "email": recipient_email, "error": str(e)}

