import { createClient } from '@supabase/supabase-js';
import { Student, Certificate, CseaEvent } from './db';

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || '';
const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || '';

if (!supabaseUrl || !supabaseAnonKey) {
  console.warn('Supabase credentials missing! Please configure NEXT_PUBLIC_SUPABASE_URL and NEXT_PUBLIC_SUPABASE_ANON_KEY in your env.');
}

export const supabase = createClient(
  supabaseUrl || 'https://placeholder.supabase.co',
  supabaseAnonKey || 'placeholder'
);

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Save Event directly to Supabase 'events' table
 */
export async function saveEventToSupabase(eventData: { event_id: string; event_name?: string; event_category?: string; event_date: string }) {
  if (!eventData || !eventData.event_id) return { success: false, error: 'No event ID' };
  try {
    const cleanId = String(eventData.event_id).trim();
    const cleanCategory = String(eventData.event_category || 'Workshop').trim();
    const cleanDate = String(eventData.event_date || new Date().toISOString().split('T')[0]).trim();

    // Upsert the event record with event_id as primary key
    const { data, error } = await supabase
      .from('events')
      .upsert([{
        event_id: cleanId,
        event_date: cleanDate,
        created_at: new Date().toISOString()
      }], { onConflict: 'event_id' })
      .select();

    if (error) {
      console.warn('Supabase saveEvent notice:', error.message);
      return { success: false, error: error.message };
    }
    return { success: true, data };
  } catch (err: any) {
    console.error('Failed to save event to Supabase:', err);
    return { success: false, error: err?.message };
  }
}

/**
 * Fetch all events directly from Supabase database 'events' table
 */
export async function fetchEventsFromSupabase(): Promise<CseaEvent[]> {
  try {
    const { data, error } = await supabase
      .from('events')
      .select('*')
      .order('created_at', { ascending: false });

    if (error || !data) return [];
    return data.map((item: any) => ({
      id: item.event_id,
      eventId: item.event_id,
      eventName: item.event_name || item.event_id,
      eventCategory: item.event_category || 'Workshop',
      eventDate: item.event_date,
      createdAt: item.created_at
    }));
  } catch (err) {
    console.warn('Supabase fetch error for events:', err);
    return [];
  }
}

/**
 * Fetch all students directly from Supabase database 'students' table
 */
export async function fetchStudentsFromSupabase(): Promise<Student[]> {
  try {
    const { data, error } = await supabase
      .from('students')
      .select('*')
      .order('created_at', { ascending: false });

    if (error || !data) {
      console.warn('Supabase fetch error for students:', error?.message);
      return [];
    }

    return data.map((item: any) => ({
      id: item.id,
      registerNo: item.register_no || item.registerNo || '',
      name: item.name || '',
      email: item.email || '',
      phone: item.phone || '',
      department: item.department || 'Computer Science and Engineering',
      yearOfStudy: item.year_of_study || '',
      section: item.section || '',
      collegeName: item.college_name || 'Kongu Engineering College',
      createdAt: item.created_at || new Date().toISOString()
    }));
  } catch (err) {
    console.error('Failed to fetch students from Supabase:', err);
    return [];
  }
}

/**
 * Fetch all certificates directly from Supabase database 'certificates' table
 */
export async function fetchCertificatesFromSupabase(): Promise<Certificate[]> {
  try {
    const { data, error } = await supabase
      .from('certificates')
      .select('*, students(*)')
      .order('created_at', { ascending: false });

    if (error || !data) {
      console.warn('Supabase fetch error for certificates:', error?.message);
      return [];
    }

    return data.map((item: any) => {
      const st = item.students || {};
      const cf = item.custom_fields || {};
      const fetchedEventName = item.event_name
        || cf.Event 
        || cf.event 
        || cf['Event Name']
        || cf.event_name 
        || item.event_id 
        || 'Workshop';
      const fetchedTitle = item.title
        || cf.Title
        || cf.title
        || cf['Paper Title']
        || cf['Project Title']
        || '';
      const fetchedEventDate = item.event_date 
        || cf.event_date 
        || '';

      const sName = st.name || item.student_name || 'Participant';
      const sRoll = st.register_no || item.student_id || '';
      const sEmail = item.student_email || st.email || '';

      return {
        id: item.id,
        certificateCode: item.id ? item.id.substring(0, 8).toUpperCase() : 'CERT-001',
        studentId: sRoll,
        studentName: sName,
        studentEmail: sEmail,
        eventId: item.event_id || '',
        eventName: fetchedEventName,
        eventDate: fetchedEventDate,
        customFields: {
          ...st,
          ...cf,
          Event: fetchedEventName,
          event_name: fetchedEventName,
          Title: fetchedTitle,
          title: fetchedTitle,
          event_date: fetchedEventDate
        },
        issueDate: item.issue_date || '',
        emailStatus: item.email_status || 'pending',
        emailSentAt: item.email_sent_at || '',
        createdAt: item.created_at || new Date().toISOString()
      };
    });
  } catch (err) {
    console.error('Failed to fetch certificates from Supabase:', err);
    return [];
  }
}

/**
 * Upsert students directly into Supabase database 'students' table and return created/updated records
 */
export async function saveStudentsToSupabase(students: Student[]) {
  if (!students || students.length === 0) return { success: false, data: [] };
  
  const records = students.map(s => ({
    register_no: String(s.registerNo || '').trim(),
    name: String(s.name || 'Participant').trim(),
    email: String(s.email || '').trim(),
    phone: String(s.phone || '').trim(),
    department: String(s.department || 'Computer Science and Engineering').trim(),
    year_of_study: String(s.yearOfStudy || '').trim(),
    section: String(s.section || '').trim(),
    college_name: String(s.collegeName || 'Kongu Engineering College').trim(),
    created_at: s.createdAt || new Date().toISOString()
  }));

  try {
    const { data, error } = await supabase
      .from('students')
      .upsert(records, { onConflict: 'register_no' })
      .select();

    return { success: true, data: data || [] };
  } catch (err: any) {
    console.error('Failed to save students to Supabase:', err);
    return { success: false, error: err?.message, data: [] };
  }
}

/**
 * Directly update a single student record in Supabase database 'students' table
 */
export async function updateStudentInSupabase(student: Student) {
  if (!student) return { success: false, error: 'No student data' };

  const payload = {
    register_no: String(student.registerNo || '').trim(),
    name: String(student.name || '').trim(),
    email: String(student.email || '').trim(),
    phone: String(student.phone || '').trim(),
    department: String(student.department || 'Computer Science and Engineering').trim(),
    year_of_study: String(student.yearOfStudy || '').trim(),
    section: String(student.section || '').trim(),
    college_name: String(student.collegeName || 'Kongu Engineering College').trim()
  };

  try {
    let res;
    if (student.id && UUID_REGEX.test(student.id)) {
      res = await supabase.from('students').update(payload).eq('id', student.id).select();
    } else {
      res = await supabase.from('students').upsert([payload], { onConflict: 'register_no' }).select();
    }

    if (res.error) {
      console.error('Error updating student in Supabase:', res.error.message);
      return { success: false, error: res.error.message };
    }
    return { success: true, data: res.data };
  } catch (err: any) {
    console.error('Failed to update student in Supabase:', err);
    return { success: false, error: err?.message };
  }
}

/**
 * Fetch App Settings (Email Template Customizations) from Supabase DB
 */
export async function fetchAppSettingsFromSupabase() {
  try {
    const { data, error } = await supabase.from('app_settings').select('*');
    if (error || !data) return {};
    const settings: Record<string, any> = {};
    data.forEach((row: any) => {
      settings[row.key] = row.value;
    });
    return settings;
  } catch (err) {
    console.error('Error fetching app_settings from Supabase:', err);
    return {};
  }
}

/**
 * Save App Settings to Supabase DB
 */
export async function saveAppSettingsToSupabase(key: string, value: any) {
  try {
    const { error } = await supabase
      .from('app_settings')
      .upsert({ key, value, updated_at: new Date().toISOString() }, { onConflict: 'key' });
    if (error) throw error;
    return { success: true };
  } catch (err: any) {
    console.error('Error saving app_settings to Supabase:', err);
    return { success: false, error: err?.message };
  }
}

/**
 * Upload an image asset to Supabase Storage 'templates' bucket
 */
export async function uploadAssetToSupabase(file: File, folder: string = 'assets'): Promise<{ success: boolean; url?: string; error?: string }> {
  try {
    const cleanFileName = `${Date.now()}_${file.name.replace(/[^a-zA-Z0-9._-]/g, '_')}`;
    const filePath = `${folder}/${cleanFileName}`;

    const { data, error } = await supabase.storage
      .from('templates')
      .upload(filePath, file, {
        cacheControl: '3600',
        upsert: true,
        contentType: file.type || 'image/png'
      });

    if (error) throw error;

    const { data: publicUrlData } = supabase.storage
      .from('templates')
      .getPublicUrl(filePath);

    return { success: true, url: publicUrlData.publicUrl };
  } catch (err: any) {
    console.warn('Supabase storage upload fallback note:', err);
    return { success: false, error: err?.message || 'Storage upload error' };
  }
}

/**
 * Upsert certificates directly into Supabase database 'certificates' table
 */
export async function saveCertificatesToSupabase(certs: Certificate[], dbStudentsOverride?: any[]) {
  if (!certs || certs.length === 0) return { success: false };

  // Pre-register any unique event IDs into 'events' table to satisfy relationships
  try {
    const uniqueMap = new Map<string, string>();
    certs.forEach(c => {
      const eid = (c.eventId || '').trim();
      if (eid && !uniqueMap.has(eid)) {
        uniqueMap.set(eid, eid);
      }
    });

    for (const [eid] of Array.from(uniqueMap.entries())) {
      await supabase.from('events').upsert({
        event_id: eid,
        event_date: certs[0]?.eventDate || new Date().toISOString().split('T')[0],
        created_at: new Date().toISOString()
      }, { onConflict: 'event_id' });
    }
  } catch (e) {
    console.warn('Pre-registering events note:', e);
  }

  // Build lookup map for students (register_no -> student_uuid)
  const studentMap: Record<string, string> = {};

  const dbStudents = (dbStudentsOverride && dbStudentsOverride.length > 0) 
    ? dbStudentsOverride 
    : await fetchStudentsFromSupabase();

  dbStudents.forEach(s => {
    const reg = (s.register_no || s.registerNo || '').toString().trim().toLowerCase();
    const sid = s.id || s.student_id;
    if (reg && sid) {
      studentMap[reg] = sid;
    }
  });

  const records = certs.map(c => {
    let validStudentUuid: string | null = null;
    const rawStId = String(c.studentId || '').trim();

    if (UUID_REGEX.test(rawStId)) {
      validStudentUuid = rawStId;
    } else if (studentMap[rawStId.toLowerCase()]) {
      validStudentUuid = studentMap[rawStId.toLowerCase()];
    }

    const effectiveEvent = c.eventName || c.eventId || '';
    const effectiveTitle = c.customFields?.Title 
      || c.customFields?.title 
      || c.customFields?.['Paper Title'] 
      || c.customFields?.['Project Title'] 
      || '';

    return {
      student_id: validStudentUuid,
      event_id: c.eventId || effectiveEvent,
      student_name: c.studentName,
      student_email: c.studentEmail || '',
      event_name: c.eventName || effectiveEvent,
      title: effectiveTitle,
      issue_date: c.issueDate || new Date().toISOString().split('T')[0],
      email_status: c.emailStatus || 'pending',
      created_at: c.createdAt || new Date().toISOString()
    };
  });

  try {
    for (const rec of records) {
      if (rec.student_id && rec.event_id) {
        await supabase
          .from('certificates')
          .delete()
          .eq('student_id', rec.student_id)
          .eq('event_id', rec.event_id);
      }
    }

    const { data, error } = await supabase
      .from('certificates')
      .insert(records)
      .select();

    if (error) {
      console.error('Error inserting certificates into Supabase:', error.message);
      return { success: false, error: error.message };
    } else {
      console.log(`[Supabase] Successfully saved ${data?.length || records.length} certificates to Supabase DB!`);
      return { success: true, count: data?.length || records.length };
    }
  } catch (err: any) {
    console.error('Failed to save certificates to Supabase:', err);
    return { success: false, error: err?.message };
  }
}

/**
 * Upload PPTX template file to Supabase Storage bucket 'templates/{event_id}/{filename}'
 */
export async function uploadTemplateToSupabaseStorage(file: File, eventId: string) {
  if (!file || !eventId) return { success: false, error: 'File or Event ID missing' };
  try {
    const filePath = `${eventId}/${file.name}`;
    
    // Upload file to Supabase Storage bucket 'templates'
    const { data, error } = await supabase.storage
      .from('templates')
      .upload(filePath, file, {
        cacheControl: '3600',
        upsert: true
      });

    if (error) {
      console.warn('Supabase storage upload notice:', error.message);
      return { success: false, error: error.message };
    }

    const { data: publicUrlData } = supabase.storage
      .from('templates')
      .getPublicUrl(filePath);

    return { success: true, path: filePath, url: publicUrlData?.publicUrl };
  } catch (err: any) {
    console.error('Failed to upload template to Supabase Storage:', err);
    return { success: false, error: err?.message };
  }
}
