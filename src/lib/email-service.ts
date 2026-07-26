export interface EmailPayload {
  recipientEmail: string;
  recipientName: string;
  eventName: string;
  certificateId: string;
  issueDate: string;
  logoUrl?: string;
  heroUrl?: string;
}

export function buildBrandedEmailHtml(payload: EmailPayload): string {
  const logo = payload.logoUrl || '/csea_logo.png';
  const hero = payload.heroUrl || '/hero.png';

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f3f4f6; margin: 0; padding: 20px; color: #1f2937; }
    .container { max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 10px 25px rgba(0,0,0,0.08); }
    .header { background: #082849; padding: 24px; text-align: center; border-bottom: 3px solid #eab308; }
    .header img { max-height: 70px; width: auto; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.2)); }
    .hero-container { width: 100%; height: 220px; overflow: hidden; background: #1e293b; }
    .hero-container img { width: 100%; height: 100%; object-fit: cover; }
    .body-content { padding: 32px 28px; }
    .title { font-size: 22px; font-weight: 700; color: #0c3f6e; margin-top: 0; margin-bottom: 12px; }
    .recipient-name { color: #026fc7; font-weight: 700; }
    .event-badge { display: inline-block; background: #f0f7ff; color: #0358a1; font-weight: 600; padding: 6px 14px; border-radius: 20px; border: 1px solid #bae0fd; margin: 8px 0 16px 0; }
    .cert-box { background: #f8fafc; border-left: 4px solid #0c8ce9; padding: 16px; margin: 20px 0; border-radius: 0 8px 8px 0; }
    .cert-label { font-size: 12px; text-transform: uppercase; color: #64748b; font-weight: 600; letter-spacing: 0.5px; }
    .cert-id { font-size: 16px; font-family: monospace; font-weight: 700; color: #0f172a; margin-top: 4px; }
    .footer { background: #f1f5f9; padding: 20px; text-align: center; font-size: 13px; color: #64748b; border-top: 1px solid #e2e8f0; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <img src="${logo}" alt="CSE Association Logo" />
    </div>
    
    <div class="hero-container">
      <img src="${hero}" alt="CSEA Hero Banner" />
    </div>

    <div class="body-content">
      <h2 class="title">🎉 Congratulations, <span class="recipient-name">${payload.recipientName}</span>!</h2>
      <p>We are delighted to present your official Certificate of Participation for:</p>
      
      <div class="event-badge">
        🎓 ${payload.eventName}
      </div>

      <p>Organized by the <strong>Computer Science Engineering Association (CSEA)</strong>. Your dedication and participation are greatly appreciated!</p>

      <div class="cert-box">
        <div style="font-size: 14px; color: #1e293b; font-weight: 600;">Issue Date: ${payload.issueDate}</div>
      </div>

      <p style="font-size: 14px; color: #475569;">📎 Your official PDF certificate is attached to this email. You can also view and download it anytime from your student profile.</p>
    </div>

    <div class="footer">
      <strong>Computer Science Engineering Association</strong><br>
      Department of Computer Science and Engineering<br>
      © ${new Date().getFullYear()} CSEA. All rights reserved.
    </div>
  </div>
</body>
</html>
  `.trim();
}
