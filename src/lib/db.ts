export interface Student {
  id: string;
  registerNo: string;
  name: string;
  email: string;
  phone?: string;
  department: string;
  yearOfStudy: string;
  section: string;
  collegeName: string;
  createdAt: string;
}

export interface CseaEvent {
  id?: string;
  eventId: string;
  eventName: string;
  eventCategory: string;
  eventDate: string;
  createdAt?: string;
  description?: string;
  htmlTemplate?: string;
  cssStyles?: string;
  excelColumnMap?: Record<string, string>;
  status?: string;
  totalParticipants?: number;
}

export interface Certificate {
  id: string;
  certificateCode: string;
  studentId: string;
  studentName: string;
  studentEmail: string;
  eventId: string;
  eventName: string;
  customFields: Record<string, any>;
  issueDate: string;
  emailStatus: 'pending' | 'sent' | 'delivered' | 'failed';
  emailSentAt?: string;
  pdfUrl?: string;
  createdAt: string;
}

export interface AssociationSettings {
  associationName: string;
  logoUrl: string;
  heroImageUrl: string;
  updatedAt: string;
}

export const DEFAULT_SETTINGS: AssociationSettings = {
  associationName: 'Computer Science Engineering Association',
  logoUrl: '/csea_logo.png',
  heroImageUrl: '/hero.png',
  updatedAt: new Date().toISOString()
};

export const DEFAULT_CERTIFICATE_TEMPLATE = `
<div style="width: 100%; height: 100%; box-sizing: border-box; padding: 30px; background: #ffffff; position: relative; font-family: 'Merriweather', 'Georgia', serif; border: 12px solid #082849;">
  <div style="width: 100%; height: 100%; border: 3px solid #eab308; padding: 25px; box-sizing: border-box; position: relative; background: radial-gradient(circle at center, #ffffff 0%, #f8fafc 100%);">
    
    <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #cbd5e1; padding-bottom: 15px; margin-bottom: 25px;">
      <div style="display: flex; align-items: center; gap: 15px;">
        <img src="/csea_logo.png" alt="CSEA Logo" style="height: 65px; width: auto;" />
        <div>
          <h3 style="margin: 0; font-size: 18px; color: #082849; font-weight: 700; letter-spacing: 0.5px;">COMPUTER SCIENCE ENGINEERING ASSOCIATION</h3>
          <p style="margin: 2px 0 0 0; font-size: 13px; color: #64748b; font-family: sans-serif;">Department of Computer Science and Engineering</p>
        </div>
      </div>
      <div style="text-align: right;">
        <span style="display: inline-block; background: #082849; color: #ffffff; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 4px; font-family: sans-serif; letter-spacing: 1px;">OFFICIAL CERTIFICATE</span>
      </div>
    </div>

    <div style="text-align: center; margin-top: 15px;">
      <h1 style="font-size: 34px; color: #082849; margin: 0; font-weight: 800; font-family: 'Georgia', serif; letter-spacing: 2px; text-transform: uppercase;">Certificate of Participation</h1>
      <p style="font-size: 15px; color: #ca8a04; font-style: italic; margin-top: 6px; font-weight: 600;">PROUDLY PRESENTED TO</p>
    </div>

    <div style="text-align: center; margin: 20px 0;">
      <h2 style="font-size: 32px; color: #0c8ce9; margin: 0; font-weight: 700; border-bottom: 2px stroke #0c8ce9; display: inline-block; padding: 0 20px 5px 20px;">
        <<Name>>
      </h2>
    </div>

    <div style="text-align: center; max-width: 85%; margin: 0 auto; line-height: 1.7; color: #334155; font-size: 15px; font-family: sans-serif;">
      <p style="margin: 0;">
        holding Roll Number <strong><<Roll Number >></strong> of Section <strong><<Section>></strong> for actively participating in the workshop on
      </p>
      <h3 style="font-size: 20px; color: #082849; margin: 8px 0; font-weight: 700; text-transform: capitalize;">
        <<event_name>>
      </h3>
      <p style="margin: 0; font-size: 14px; color: #475569;">
        organized by the Computer Science Engineering Association (CSEA), <<college_name>>.
      </p>
    </div>

  </div>
</div>
`;
