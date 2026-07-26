import { CseaEvent, Certificate, Student, DEFAULT_CERTIFICATE_TEMPLATE } from './db';

const INITIAL_EVENT_ID = 'evt-ml-workshop-2026';

export const INITIAL_EVENT: CseaEvent = {
  id: INITIAL_EVENT_ID,
  eventId: INITIAL_EVENT_ID,
  eventName: 'Machine Learning Workshop',
  eventCategory: 'workshop',
  eventDate: '2026-07-23',
  description: 'Hands-on Workshop covering AI/ML fundamentals, Python model training, and project development.',
  htmlTemplate: DEFAULT_CERTIFICATE_TEMPLATE,
  cssStyles: '',
  excelColumnMap: {
    'Name': 'Name',
    'Roll Number ': 'Roll Number ',
    'Mail id (College mail id)': 'Mail id (College mail id)',
    'Section': 'Section'
  },
  status: 'completed',
  createdAt: '2026-07-23T10:00:00Z',
  totalParticipants: 0
};

// Store manager helpers (works in SSR and Client)
export class StoreManager {
  private static STORAGE_KEY_EVENTS = 'csea_events_data_v3';
  private static STORAGE_KEY_CERTS = 'csea_certs_data_v3';
  private static STORAGE_KEY_STUDENTS = 'csea_students_data_v3';

  static getEvents(): CseaEvent[] {
    if (typeof window === 'undefined') return [INITIAL_EVENT];
    const saved = localStorage.getItem(this.STORAGE_KEY_EVENTS);
    if (!saved) {
      localStorage.setItem(this.STORAGE_KEY_EVENTS, JSON.stringify([INITIAL_EVENT]));
      return [INITIAL_EVENT];
    }
    try {
      return JSON.parse(saved);
    } catch {
      return [INITIAL_EVENT];
    }
  }

  static getCertificates(): Certificate[] {
    if (typeof window === 'undefined') return [];
    const saved = localStorage.getItem(this.STORAGE_KEY_CERTS);
    if (!saved) return [];
    try {
      return JSON.parse(saved);
    } catch {
      return [];
    }
  }

  static getStudents(): Student[] {
    if (typeof window === 'undefined') return [];
    const saved = localStorage.getItem(this.STORAGE_KEY_STUDENTS);
    if (!saved) return [];
    try {
      return JSON.parse(saved);
    } catch {
      return [];
    }
  }

  static saveCertificates(certs: Certificate[]): void {
    if (typeof window !== 'undefined') {
      localStorage.setItem(this.STORAGE_KEY_CERTS, JSON.stringify(certs));
    }
  }

  static saveStudents(students: Student[]): void {
    if (typeof window !== 'undefined') {
      localStorage.setItem(this.STORAGE_KEY_STUDENTS, JSON.stringify(students));
    }
  }
}
