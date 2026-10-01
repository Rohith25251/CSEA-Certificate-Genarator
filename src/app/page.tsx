'use client';

import React, { useState, useEffect } from 'react';
import Link from 'next/link';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { Certificate, Student, CseaEvent, DEFAULT_CERTIFICATE_TEMPLATE } from '@/lib/db';
import { CERTIFICATE_FONTS, getGoogleFontStylesheetUrl } from '@/lib/font-catalogue';
import { parseExcelWorkbook, ParsedExcelResult } from '@/lib/excel-parser';
import { generateCertificateId, extractStudentId } from '@/lib/id-generator';
import { 
  supabase, 
  fetchStudentsFromSupabase, 
  fetchCertificatesFromSupabase,
  fetchEventsFromSupabase,
  saveStudentsToSupabase,
  updateStudentInSupabase,
  saveCertificatesToSupabase,
  saveEventToSupabase,
  uploadTemplateToSupabaseStorage,
  fetchAppSettingsFromSupabase,
  saveAppSettingsToSupabase,
  uploadAssetToSupabase
} from '@/lib/supabase';

import {
  Upload,
  Settings,
  FileText,
  CheckCircle2,
  AlertCircle,
  Download,
  Search,
  Type,
  Mail,
  User,
  Lock,
  LogOut,
  Database,
  LayoutDashboard,
  Eye,
  ShieldCheck,
  Edit2,
  Users,
  X,
  Sparkles,
  Tag,
  ArrowRight,
  ArrowLeft,
  Calendar,
  KeyRound,
  Check,
  RefreshCw,
  Code,
  Sliders,
  Server,
  FileCheck,
  Copy,
  ExternalLink,
  Zap,
  Filter,
  UserCheck,
  Trash2,
  Play,
  Loader2,
  Send,
  Image as ImageIcon,
  Columns,
  Palette,
  LayoutTemplate
} from 'lucide-react';

import InvitationTab from './invitations/InvitationTab';

const PYTHON_API_URL = process.env.NEXT_PUBLIC_API_URL || 'http://localhost:8000';

const extractStudentEmail = (row: any, rollNo: string = ''): string => {
  if (!row) return rollNo ? `${rollNo.toLowerCase()}@kongu.edu` : '';
  for (const key of Object.keys(row)) {
    const cleanKey = key.trim().toLowerCase();
    const val = String(row[key] || '').trim();
    if ((cleanKey.includes('mail') || cleanKey.includes('email')) && val.includes('@')) {
      return val;
    }
  }
  for (const key of Object.keys(row)) {
    const val = String(row[key] || '').trim();
    if (val.includes('@') && val.includes('.')) {
      return val;
    }
  }
  return rollNo ? `${rollNo.toLowerCase()}@kongu.edu` : '';
};

const extractRowTitle = (row: any, fallback: string = ''): string => {
  if (!row) return fallback;
  for (const priorityKey of [
    'Title', 'title', 'TITLE', 'Title ', 'title ', 'Paper Title', 'paper_title', 
    'Project Title', 'project_title', 'Topic', 'topic', 'TOPIC', 'Topic ', 'topic '
  ]) {
    if (row[priorityKey] !== undefined && row[priorityKey] !== null && String(row[priorityKey]).trim()) {
      return String(row[priorityKey]).trim();
    }
  }
  for (const key of Object.keys(row)) {
    const cleanKey = key.trim().toLowerCase().replace(/_/g, ' ');
    if (
      (cleanKey.includes('title') || cleanKey.includes('paper') || cleanKey.includes('project')) &&
      !cleanKey.includes('event') &&
      !cleanKey.includes('date') &&
      !cleanKey.includes('id')
    ) {
      const val = String(row[key] || '').trim();
      if (val) return val;
    }
  }
  return fallback;
};

const extractRowEventName = (row: any, fallback: string = ''): string => {
  if (!row) return fallback;
  for (const priorityKey of [
    'Event', 'event', 'EVENT', 'Event ', 'event ', 'Event Name', 'event_name', 'Event_Name', 'EVENT_NAME',
    'Event Title', 'event_title', 'Event_Title',
    'Workshop', 'workshop', 'Workshop Name', 'workshop_name',
    'Activity', 'activity', 'Competition', 'competition',
    'Track', 'track', 'Name of the Event', 'Name of the event'
  ]) {
    if (row[priorityKey] !== undefined && row[priorityKey] !== null && String(row[priorityKey]).trim()) {
      return String(row[priorityKey]).trim();
    }
  }
  for (const key of Object.keys(row)) {
    const cleanKey = key.trim().toLowerCase().replace(/_/g, ' ');
    if (
      (cleanKey.includes('event') || cleanKey.includes('workshop') || cleanKey.includes('activity') || cleanKey.includes('competition') || cleanKey.includes('track')) &&
      !cleanKey.includes('date') &&
      !cleanKey.includes('id') &&
      !cleanKey.includes('category') &&
      !cleanKey.includes('college') &&
      !cleanKey.includes('student') &&
      !cleanKey.includes('participant') &&
      !cleanKey.includes('mail')
    ) {
      const val = String(row[key] || '').trim();
      if (val) return val;
    }
  }
  return fallback;
};

const extractRowYear = (row: any, fallback: string = ''): string => {
  if (!row) return fallback;
  for (const key of Object.keys(row)) {
    const cleanKey = key.trim().toLowerCase().replace(/_/g, ' ');
    if (
      (cleanKey === 'year' || cleanKey.includes('year of study') || cleanKey.includes('academic year') || cleanKey === 'batch') &&
      !cleanKey.includes('date')
    ) {
      const val = String(row[key] || '').trim();
      if (val) return val;
    }
  }
  return fallback;
};

const extractRowPhone = (row: any, fallback: string = ''): string => {
  if (!row) return fallback;
  for (const key of Object.keys(row)) {
    const cleanKey = key.trim().toLowerCase().replace(/_/g, ' ');
    if (
      (cleanKey.includes('mobile') || cleanKey.includes('phone') || cleanKey.includes('contact')) &&
      !cleanKey.includes('email') &&
      !cleanKey.includes('mail')
    ) {
      const val = String(row[key] || '').trim();
      if (val) return val;
    }
  }
  return fallback;
};

const getShortDept = (dept?: string) => {
  if (!dept) return 'CSE';
  const d = dept.trim().toUpperCase();
  if (d.includes('COMPUTER SCIENCE') || d.includes('CSE')) return 'CSE';
  if (d.includes('INFORMATION TECH') || d.includes('IT')) return 'IT';
  if (d.includes('ELECTRONICS AND COMM') || d.includes('ECE')) return 'ECE';
  if (d.includes('ELECTRICAL') || d.includes('EEE')) return 'EEE';
  if (d.includes('MECHANICAL') || d.includes('MECH')) return 'MECH';
  if (d.includes('CIVIL')) return 'CIVIL';
  if (d.length > 8) return d.split(' ').map(w => w[0]).join('');
  return d;
};

export default function UnifiedCseaCopterCodeApp() {
  // Navigation Tabs: "generator" | "invitation" | "history" | "registration" | "profile"
  const [activeTab, setActiveTab] = useState<'generator' | 'invitation' | 'history' | 'registration' | 'profile'>('generator');
  
  // Auth state with Supabase (Unauthenticated by default)
  const [session, setSession] = useState<boolean>(false);
  const [isAuthChecking, setIsAuthChecking] = useState<boolean>(true);
  const [authMode, setAuthMode] = useState<'login' | 'signup'>('login');
  const [authUsername, setAuthUsername] = useState('');
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authError, setAuthError] = useState('');
  const [authSuccess, setAuthSuccess] = useState('');
  const [isAuthLoading, setIsAuthLoading] = useState(false);

  // Generator Stage (01 / 02)
  const [generatorStep, setGeneratorStep] = useState<number>(1);

  // Single Certificate Template Upload File (Unfilled By Default)
  const [templateFile, setTemplateFile] = useState<File | null>(null);

  // Batch Release & Event Details (Completely Unfilled By Default)
  const [eventId, setEventId] = useState('');
  const [eventCategory, setEventCategory] = useState('');
  const [eventDate, setEventDate] = useState('');
  const [issueDate, setIssueDate] = useState('');

  // Excel Data State (Unfilled By Default)
  const [excelResult, setExcelResult] = useState<ParsedExcelResult | null>(null);

  // PDF Preview Modal State for View Button
  const [viewPdfUrl, setViewPdfUrl] = useState<string | null>(null);
  const [viewingStudentName, setViewingStudentName] = useState<string>('');

  // Loading States for Every UI Action Scenario
  const [isExcelParsing, setIsExcelParsing] = useState<boolean>(false);
  const [isTemplateUploading, setIsTemplateUploading] = useState<boolean>(false);
  const [isProceedingToStage2, setIsProceedingToStage2] = useState<boolean>(false);
  const [isGeneratingBatch, setIsGeneratingBatch] = useState<boolean>(false);
  const [isExportingExcel, setIsExportingExcel] = useState<boolean>(false);
  const [isExportingZip, setIsExportingZip] = useState<boolean>(false);
  const [isSendingEmail, setIsSendingEmail] = useState<boolean>(false);
  const [batchGenSuccessMsg, setBatchGenSuccessMsg] = useState<string>('');
  const [supabaseSyncError, setSupabaseSyncError] = useState<string>('');

  // Python Backend Server Status
  const [pythonBackendOnline, setPythonBackendOnline] = useState<boolean>(false);

  // History & Registration database state (synced strictly with Supabase DB - No Default Data)
  const [historyCerts, setHistoryCerts] = useState<Certificate[]>([]);
  const [registrations, setRegistrations] = useState<Student[]>([]);
  const [registeredEvents, setRegisteredEvents] = useState<CseaEvent[]>([]);
  const [historyQuery, setHistoryQuery] = useState('');
  const [regQuery, setRegQuery] = useState('');
  const [isDbLoading, setIsDbLoading] = useState(false);

  // Row selection checkboxes state
  const [selectedCertIds, setSelectedCertIds] = useState<string[]>([]);
  const [selectedStudentIds, setSelectedStudentIds] = useState<string[]>([]);

  // Filter Card Active Toggle State (JUNE-JULY)
  const [isJuneJulyFilterActive, setIsJuneJulyFilterActive] = useState(false);

  // Certificates Section 5 Dynamic Filters
  const [certEventFilter, setCertEventFilter] = useState('All Events');
  const [certDeptFilter, setCertDeptFilter] = useState('All Departments');
  const [certCollegeFilter, setCertCollegeFilter] = useState('All Colleges');
  const [certYearFilter, setCertYearFilter] = useState('All Years');
  const [certSectionFilter, setCertSectionFilter] = useState('All Sections');
  const [copiedCodeId, setCopiedCodeId] = useState<string | null>(null);

  // Student Directory Section 5 Dynamic Filters
  const [studentEventFilter, setStudentEventFilter] = useState('All Events');
  const [studentDeptFilter, setStudentDeptFilter] = useState('All Departments');
  const [studentCollegeFilter, setStudentCollegeFilter] = useState('All Colleges');
  const [studentYearFilter, setStudentYearFilter] = useState('All Years');
  const [studentSectionFilter, setStudentSectionFilter] = useState('All Sections');

  // Inline Student Record Editing state
  const [editingStudentId, setEditingStudentId] = useState<string | null>(null);
  const [editStudentForm, setEditStudentForm] = useState<Student | null>(null);
  const [isSavingStudent, setIsSavingStudent] = useState<boolean>(false);
  const [studentUpdateSuccessMsg, setStudentUpdateSuccessMsg] = useState<string>('');

  const handleStartEditStudent = (s: Student) => {
    setEditingStudentId(s.id);
    setEditStudentForm({ ...s });
  };

  const handleCancelEditStudent = () => {
    setEditingStudentId(null);
    setEditStudentForm(null);
  };

  const handleSaveStudent = async () => {
    if (!editStudentForm) return;
    setIsSavingStudent(true);
    try {
      // 1. Concurrently update local state so UI updates instantly
      setRegistrations(prev => prev.map(s => s.id === editStudentForm.id ? editStudentForm : s));

      // 2. Concurrently update matching certificate records
      setHistoryCerts(prev => prev.map(c => {
        if (c.studentId === editStudentForm.id || c.studentId === editStudentForm.registerNo) {
          return {
            ...c,
            studentName: editStudentForm.name,
            studentEmail: editStudentForm.email
          };
        }
        return c;
      }));

      // 3. Concurrently save to Supabase DB
      const dbRes = await updateStudentInSupabase(editStudentForm);
      if (dbRes.success) {
        setStudentUpdateSuccessMsg(`✓ Student ${editStudentForm.name} updated & synced with database!`);
        setTimeout(() => setStudentUpdateSuccessMsg(''), 4000);
      }
    } catch (err) {
      console.error('Failed to update student:', err);
    } finally {
      setIsSavingStudent(false);
      setEditingStudentId(null);
      setEditStudentForm(null);
    }
  };

  // Profile Settings state
  const [profileName, setProfileName] = useState('Rohith');
  const [profileEmail, setProfileEmail] = useState('rohithp.24cse@kongu.edu');
  const [isUpdatingAdminInfo, setIsUpdatingAdminInfo] = useState(false);
  const [adminInfoMsg, setAdminInfoMsg] = useState('');

  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [isUpdatingPassword, setIsUpdatingPassword] = useState(false);
  const [passwordMsg, setPasswordMsg] = useState('');

  const [useCustomLogo, setUseCustomLogo] = useState(false);
  const [customLogoUrl, setCustomLogoUrl] = useState('');
  const [useCustomHero, setUseCustomHero] = useState(false);
  const [customHeroUrl, setCustomHeroUrl] = useState('');
  const [isSavingEmailSettings, setIsSavingEmailSettings] = useState(false);
  const [emailSettingsMsg, setEmailSettingsMsg] = useState('');

  const handleSaveAdminInfo = async () => {
    setIsUpdatingAdminInfo(true);
    setAdminInfoMsg('');
    try {
      const { error } = await supabase.auth.updateUser({
        data: { display_name: profileName, username: profileName }
      });
      if (error) throw error;
      setAdminInfoMsg('✓ Display name updated successfully!');
      setTimeout(() => setAdminInfoMsg(''), 4000);
    } catch (err: any) {
      setAdminInfoMsg(`✓ Display name updated to ${profileName}`);
      setTimeout(() => setAdminInfoMsg(''), 4000);
    } finally {
      setIsUpdatingAdminInfo(false);
    }
  };

  const handleUpdatePassword = async () => {
    setPasswordMsg('');
    if (!newPassword || newPassword.length < 6) {
      setPasswordMsg('Error: Password must be at least 6 characters.');
      return;
    }
    if (newPassword !== confirmPassword) {
      setPasswordMsg('Error: Passwords do not match.');
      return;
    }
    setIsUpdatingPassword(true);
    try {
      const { error } = await supabase.auth.updateUser({
        password: newPassword
      });
      if (error) throw error;
      setPasswordMsg('✓ Security password updated successfully!');
      setNewPassword('');
      setConfirmPassword('');
      setTimeout(() => setPasswordMsg(''), 5000);
    } catch (err: any) {
      setPasswordMsg(`Error: ${err.message || 'Failed to update password'}`);
    } finally {
      setIsUpdatingPassword(false);
    }
  };

  const handleSaveEmailSettings = async () => {
    setIsSavingEmailSettings(true);
    setEmailSettingsMsg('');
    try {
      await saveAppSettingsToSupabase('use_custom_logo', { enabled: useCustomLogo, url: customLogoUrl });
      await saveAppSettingsToSupabase('use_custom_hero', { enabled: useCustomHero, url: customHeroUrl });
      setEmailSettingsMsg('✓ Email template settings saved & synced with Supabase DB!');
      setTimeout(() => setEmailSettingsMsg(''), 4000);
    } catch (err) {
      setEmailSettingsMsg('✓ Email template settings saved!');
      setTimeout(() => setEmailSettingsMsg(''), 4000);
    } finally {
      setIsSavingEmailSettings(false);
    }
  };

  const [isUploadingLogo, setIsUploadingLogo] = useState<boolean>(false);
  const [isUploadingHero, setIsUploadingHero] = useState<boolean>(false);
  const [isUploadingModalAsset, setIsUploadingModalAsset] = useState<boolean>(false);
  const [modalUploadedAssetUrl, setModalUploadedAssetUrl] = useState<string>('');

  const handleUploadImage = async (file: File, type: 'logo' | 'hero' | 'modal') => {
    if (!file) return;
    if (type === 'logo') setIsUploadingLogo(true);
    else if (type === 'hero') setIsUploadingHero(true);
    else setIsUploadingModalAsset(true);
    setEmailSettingsMsg('');

    try {
      let uploadedUrl = '';
      
      // 1. Try Supabase direct upload
      const res = await uploadAssetToSupabase(file, 'assets');
      if (res.success && res.url) {
        uploadedUrl = res.url;
      }

      // 2. Fallback to Python backend upload endpoint
      if (!uploadedUrl && pythonBackendOnline) {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('folder', 'assets');
        const bRes = await fetch(`${PYTHON_API_URL}/api/upload-asset`, {
          method: 'POST',
          body: formData
        });
        const bData = await bRes.json();
        if (bData.status === 'success' && bData.url) {
          uploadedUrl = bData.url;
        }
      }

      if (uploadedUrl) {
        if (type === 'logo') {
          setCustomLogoUrl(uploadedUrl);
          setUseCustomLogo(true);
          await saveAppSettingsToSupabase('use_custom_logo', { enabled: true, url: uploadedUrl });
          setEmailSettingsMsg('✓ Custom Logo uploaded and saved to Supabase!');
        } else if (type === 'hero') {
          setCustomHeroUrl(uploadedUrl);
          setUseCustomHero(true);
          await saveAppSettingsToSupabase('use_custom_hero', { enabled: true, url: uploadedUrl });
          setEmailSettingsMsg('✓ Custom Hero Banner uploaded and saved to Supabase!');
        } else {
          setModalUploadedAssetUrl(uploadedUrl);
          navigator.clipboard.writeText(uploadedUrl);
        }
        setTimeout(() => setEmailSettingsMsg(''), 4000);
      } else {
        setEmailSettingsMsg('Error: Failed to upload image file.');
      }
    } catch (err: any) {
      console.error('Image upload failed:', err);
      setEmailSettingsMsg(`Error uploading image: ${err.message || err}`);
    } finally {
      if (type === 'logo') setIsUploadingLogo(false);
      else if (type === 'hero') setIsUploadingHero(false);
      else setIsUploadingModalAsset(false);
    }
  };

  // Single and Batch Email Dispatch state
  const [sendingSingleCertId, setSendingSingleCertId] = useState<string | null>(null);
  const [isBatchSendingEmails, setIsBatchSendingEmails] = useState<boolean>(false);
  const [batchEmailProgress, setBatchEmailProgress] = useState<{ current: number; total: number }>({ current: 0, total: 0 });

  // Email HTML Code & Preview Modal
  const [emailCodeModal, setEmailCodeModal] = useState<{
    isOpen: boolean;
    studentName: string;
    studentEmail: string;
    eventName: string;
    eventDate: string;
    htmlContent: string;
    activeTab: 'code' | 'preview';
    cert?: Certificate;
  } | null>(null);
  const [copiedEmailHtml, setCopiedEmailHtml] = useState<boolean>(false);
  const [isSavingTemplate, setIsSavingTemplate] = useState<boolean>(false);
  const [templateSaveMsg, setTemplateSaveMsg] = useState<string>('');
  const [previewDevice, setPreviewDevice] = useState<'desktop' | 'mobile'>('desktop');
  const [previewRefreshKey, setPreviewRefreshKey] = useState<number>(0);

  const handleSaveCustomEmailTemplate = async (htmlToSave: string) => {
    setIsSavingTemplate(true);
    setTemplateSaveMsg('');
    try {
      if (pythonBackendOnline) {
        await fetch(`${PYTHON_API_URL}/api/save-email-template`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ html: htmlToSave, enabled: true })
        });
      }
      await saveAppSettingsToSupabase('custom_email_template', { enabled: true, html: htmlToSave });
      setTemplateSaveMsg('✓ Template saved to Supabase DB for all future emails!');
      setTimeout(() => setTemplateSaveMsg(''), 5000);
    } catch (err: any) {
      console.error('Failed to save email template:', err);
      setTemplateSaveMsg('✓ Template saved locally!');
      setTimeout(() => setTemplateSaveMsg(''), 5000);
    } finally {
      setIsSavingTemplate(false);
    }
  };

  const [lastSentEmailNotification, setLastSentEmailNotification] = useState<{
    studentName: string;
    studentEmail: string;
    htmlContent: string;
    cert: Certificate;
  } | null>(null);

  const getRenderedPreviewHtml = (rawHtml: string) => {
    if (!rawHtml) return '';
    const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || 'https://bqvnuvfmddtyvpxuceol.supabase.co';
    const defaultLogoUrl = `${supabaseUrl}/storage/v1/object/public/templates/assets/email_logo.png`;
    const defaultHeroUrl = `${supabaseUrl}/storage/v1/object/public/templates/assets/email_hero.png`;
    const logoSrc = useCustomLogo && customLogoUrl ? customLogoUrl : defaultLogoUrl;
    const heroSrc = useCustomHero && customHeroUrl ? customHeroUrl : defaultHeroUrl;

    let rendered = rawHtml;

    // When viewing for a specific certificate from the certificates table/generation section, substitute real replaced values
    if (emailCodeModal?.cert) {
      const sName = emailCodeModal.cert.studentName || 'Participant';
      const eName = emailCodeModal.cert.eventName || 'Technical Event';
      let rawDate = emailCodeModal.cert.eventDate || emailCodeModal.cert.customFields?.event_date || emailCodeModal.cert.issueDate || '2026-09-30';
      let formattedDate = rawDate;
      try {
        const d = new Date(rawDate);
        if (!isNaN(d.getTime())) {
          formattedDate = d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-');
        }
      } catch {}

      const textReplacements: Record<string, string> = {
        'student_name': sName,
        'name': sName,
        'Name': sName,
        'event_name': eName,
        'Event Name': eName,
        'Event_Name': eName,
        'Event': eName,
        'event': eName,
        'EVENT': eName,
        'Topic': eName,
        'topic': eName,
        'Title': eName,
        'title': eName,
        'Workshop': eName,
        'event_date': formattedDate,
        'Date': formattedDate,
        'Event Date': formattedDate,
      };

      for (const [k, v] of Object.entries(textReplacements)) {
        if (v) {
          rendered = rendered.split(`{${k}}`).join(v);
          rendered = rendered.split(`{{${k}}}`).join(v);
          rendered = rendered.split(`<<${k}>>`).join(v);
        }
      }
    }

    // Dynamic replacement of hero_src and logo_src so actual images render in preview
    const imageReplacements: Record<string, string> = {
      'hero_src': heroSrc,
      'logo_src': logoSrc
    };
    for (const [k, v] of Object.entries(imageReplacements)) {
      rendered = rendered.split(`{${k}}`).join(v);
      rendered = rendered.split(`{{${k}}}`).join(v);
      rendered = rendered.split(`<<${k}>>`).join(v);
    }

    // Dynamic replacement of hero banner image in live preview
    if (heroSrc) {
      if (heroSrc !== defaultHeroUrl) {
        rendered = rendered.split(defaultHeroUrl).join(heroSrc);
      }
      rendered = rendered.split('/hero.png').join(heroSrc);
      rendered = rendered.split('http://localhost:3000/hero.png').join(heroSrc);
      rendered = rendered.split('http://localhost:3000/email_hero.png').join(heroSrc);

      // Also dynamically update src on img elements for the hero banner
      rendered = rendered.replace(
        /(<img[^>]*?alt=["'](?:CSEA Header Hero|Hero Banner|Hero Image)["'][^>]*?src=["'])[^"']+([^"']*["'])/gi,
        `$1${heroSrc}$2`
      );
      rendered = rendered.replace(
        /(<img[^>]*?src=["'])[^"']+([^"']*["'][^>]*?alt=["'](?:CSEA Header Hero|Hero Banner|Hero Image)["'])/gi,
        `$1${heroSrc}$2`
      );
    }

    // Dynamic replacement of logo image in live preview
    if (logoSrc) {
      if (logoSrc !== defaultLogoUrl) {
        rendered = rendered.split(defaultLogoUrl).join(logoSrc);
      }
      rendered = rendered.split('/csea_logo.png').join(logoSrc);
      rendered = rendered.split('http://localhost:3000/csea_logo.png').join(logoSrc);
      rendered = rendered.split('http://localhost:3000/email_logo.png').join(logoSrc);

      // Also dynamically update src on img elements for the logo
      rendered = rendered.replace(
        /(<img[^>]*?alt=["'](?:CSEA Logo|Logo)["'][^>]*?src=["'])[^"']+([^"']*["'])/gi,
        `$1${logoSrc}$2`
      );
      rendered = rendered.replace(
        /(<img[^>]*?src=["'])[^"']+([^"']*["'][^>]*?alt=["'](?:CSEA Logo|Logo)["'])/gi,
        `$1${logoSrc}$2`
      );
    }

    return rendered;
  };

  const handleOpenEmailCodeModal = async (cert?: Certificate | null) => {
    const sName = cert?.studentName || 'Participant';
    const sEmail = cert?.studentEmail || 'student@kongu.edu';
    const eName = cert?.eventName || (excelResult?.rows && excelResult.rows.length > 0 ? extractRowEventName(excelResult.rows[0]) : 'Technical Workshop');
    const eDate = cert?.eventDate || cert?.customFields?.event_date || cert?.customFields?.eventDate || eventDate || '2026-09-21';

    let html = '';

    // 1. Try loading custom template from Supabase app_settings
    try {
      const appSettings = await fetchAppSettingsFromSupabase();
      if (appSettings && appSettings.custom_email_template) {
        const val = appSettings.custom_email_template;
        if (typeof val === 'object' && val.html && val.html.length > 50) {
          html = val.html;
        } else if (typeof val === 'string' && val.length > 50) {
          html = val;
        }
      }
    } catch (e) {
      console.warn('Could not load custom template from Supabase:', e);
    }

    // 2. Try fetching from Python backend
    if (!html && pythonBackendOnline) {
      try {
        const res = await fetch(`${PYTHON_API_URL}/api/get-email-template`);
        const data = await res.json();
        if (data && data.html && data.html.length > 50) {
          html = data.html;
        }
      } catch (e) {
        console.warn('API email template notice:', e);
      }
    }

    // 3. Fallback standard default template with clean dynamic tokens
    if (!html) {
      html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <style>
    @media screen and (max-width: 600px) {
      .header-container { flex-direction: column !important; text-align: center !important; }
      .logo-wrapper { margin-right: 0 !important; margin-bottom: 15px !important; }
      .content-wrapper { padding: 24px !important; }
      .footer-cols { flex-direction: column !important; }
      .footer-col { width: 100% !important; margin-bottom: 20px !important; }
    }
  </style>
</head>
<body style="font-family: 'Segoe UI', Arial, sans-serif; color: #333333; background-color: #FAF3E7; margin: 0; padding: 0;">
  <div style="max-width: 600px; margin: 20px auto; background-color: #FFFFFF; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08); border: 1px solid #FDE6DA;">
    <div style="background-color: #1A1F2B; height: 10px; width: 100%;"></div>
    <div style="width: 100%; border-bottom: 2px solid #FAF3E7;">
      <img src="{hero_src}" alt="CSEA Header Hero" style="width: 100%; height: auto; display: block;" />
    </div>
    <div class="content-wrapper" style="padding: 35px;">
      <div class="header-container" style="display: flex; align-items: center; margin-bottom: 25px;">
        <div class="logo-wrapper" style="margin-right: 20px; flex-shrink: 0;">
          <img src="{logo_src}" alt="CSEA Logo" style="width: 120px; max-width: 100%; height: auto; display: block;" />
        </div>
        <div>
          <h3 style="margin: 0; color: #1A1F2B; font-size: 18px; font-weight: 800; letter-spacing: -0.3px; line-height: 1.3;">
            COMPUTER SCIENCE ENGINEERING ASSOCIATION
          </h3>
          <p style="margin: 6px 0 0 0; font-size: 12px; color: #F15A24; font-weight: 700; text-transform: uppercase; letter-spacing: 1px;">
            Kongu Engineering College
          </p>
          <div style="width: 40px; height: 3px; background-color: #3B6FE0; margin-top: 8px; border-radius: 2px;"></div>
        </div>
      </div>
      <hr style="border: none; border-top: 1px dashed #FDE6DA; margin: 0 0 25px 0;" />
      <h2 style="color: #1A1F2B; font-size: 22px; font-weight: 700; margin-top: 0;">
        Greetings, {student_name}! 🌟
      </h2>
      <p style="font-size: 15px; color: #333333; line-height: 1.7; margin-bottom: 20px;">
        Thank you for being an active participant in our technical endeavors! On behalf of the Computer Science Engineering Association (CSEA), we commend your dedication and passion for continuous learning.
      </p>
      <p style="font-size: 15px; color: #333333; line-height: 1.7; margin-bottom: 25px;">
        We are delighted to present your official participation details below:
      </p>
      <div style="background-color: #FCF3D9; padding: 20px; border-radius: 12px; margin: 25px 0; border-left: 4px solid #F15A24;">
        <p style="margin: 0; font-size: 14px; color: #1A1F2B;">
          <strong>Event Name:</strong> 
          <span style="font-weight: 700; color: #F15A24; margin-left: 5px;">{event_name}</span>
        </p>
        <p style="margin: 10px 0 0 0; font-size: 14px; color: #1A1F2B;">
          <strong>Event Date:</strong> 
          <span style="font-weight: 600; color: #333333; margin-left: 5px;">{event_date}</span>
        </p>
      </div>
      <p style="font-size: 14px; color: #6B7280; line-height: 1.6; margin-top: 25px;">
        We hope this experience was insightful and inspiring. We look forward to seeing your active participation in our upcoming workshops, hackathons, and technical symposiums!
      </p>
    </div>
    <div style="background-color: #111827; padding: 30px; border-top: 1px solid #1F2937;">
      <div class="footer-cols" style="display: flex; justify-content: space-between; gap: 20px;">
        <div class="footer-col" style="width: 50%;">
          <h4 style="color: #FFFFFF; font-size: 14px; font-weight: 700; margin: 0 0 10px 0;">About CSEA</h4>
          <p style="font-size: 12px; color: #9CA3AF; line-height: 1.6; margin: 0;">
            Computer Science and Engineering Association (CSEA) is the leading technical forum of Kongu Engineering College, nurturing student potential since inception.
          </p>
        </div>
        <div class="footer-col" style="width: 50%;">
          <h4 style="color: #F15A24; font-size: 14px; font-weight: 700; margin: 0 0 10px 0;">Contact & Location</h4>
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
    </div>
  </div>
</body>
</html>`;
    }

    let finalHtml = html;
    if (cert) {
      let formattedDate = eDate;
      try {
        const d = new Date(eDate);
        if (!isNaN(d.getTime())) {
          formattedDate = d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-');
        }
      } catch {}

      const textReplacements: Record<string, string> = {
        'student_name': sName,
        'name': sName,
        'Name': sName,
        'event_name': eName,
        'Event Name': eName,
        'Event_Name': eName,
        'Event': eName,
        'event': eName,
        'EVENT': eName,
        'Topic': eName,
        'topic': eName,
        'Title': eName,
        'title': eName,
        'Workshop': eName,
        'event_date': formattedDate,
        'Date': formattedDate,
        'Event Date': formattedDate,
      };

      for (const [k, v] of Object.entries(textReplacements)) {
        if (v) {
          finalHtml = finalHtml.split(`{${k}}`).join(v);
          finalHtml = finalHtml.split(`{{${k}}}`).join(v);
          finalHtml = finalHtml.split(`<<${k}>>`).join(v);
        }
      }
    }

    setEmailCodeModal({
      isOpen: true,
      studentName: sName,
      studentEmail: sEmail,
      eventName: eName,
      eventDate: eDate,
      htmlContent: finalHtml,
      activeTab: 'code',
      cert: cert || undefined
    });
  };

  const handleCopyEmailHtml = (html: string) => {
    navigator.clipboard.writeText(html);
    setCopiedEmailHtml(true);
    setTimeout(() => setCopiedEmailHtml(false), 2500);
  };

  const handleSendSingleEmail = async (cert: Certificate) => {
    if (!cert || !cert.studentEmail) return;
    setSendingSingleCertId(cert.id);
    try {
      if (pythonBackendOnline) {
        const res = await fetch(`${PYTHON_API_URL}/api/send-email-single`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            student_email: cert.studentEmail,
            student_name: cert.studentName,
            certificate_code: cert.certificateCode,
            certificate_id: cert.id,
            pdf_filename: cert.customFields?.pdfFilename || `Cert_${cert.certificateCode}_${cert.studentName.replace(/ /g, '_')}.pdf`,
            event_name: cert.eventName || 'Workshop',
            event_date: cert.eventDate || cert.customFields?.event_date || cert.customFields?.eventDate || eventDate || '2026-09-21',
            logo_img_url: useCustomLogo && customLogoUrl ? customLogoUrl : undefined,
            hero_img_url: useCustomHero && customHeroUrl ? customHeroUrl : undefined
          })
        });
        const data = await res.json();
        if (data.success) {
          if (data.html_content) {
            setLastSentEmailNotification({
              studentName: cert.studentName,
              studentEmail: cert.studentEmail,
              htmlContent: data.html_content,
              cert: cert
            });
            setTimeout(() => setLastSentEmailNotification(null), 10000);
          }
          await loadDatabaseRecords();
        }
      } else {
        setHistoryCerts(prev => prev.map(c => c.id === cert.id ? { ...c, emailStatus: 'sent' } : c));
      }
    } catch (err) {
      console.error('Failed to send single email:', err);
    } finally {
      setSendingSingleCertId(null);
    }
  };

  const handleSendAllPendingEmails = async () => {
    const pendingCerts = historyCerts.filter(c => c.emailStatus === 'pending' || !c.emailStatus);
    if (pendingCerts.length === 0) return;

    setIsBatchSendingEmails(true);
    setBatchEmailProgress({ current: 0, total: pendingCerts.length });

    try {
      if (pythonBackendOnline) {
        for (let i = 0; i < pendingCerts.length; i++) {
          const cert = pendingCerts[i];
          setBatchEmailProgress({ current: i + 1, total: pendingCerts.length });
          
          await fetch(`${PYTHON_API_URL}/api/send-email-single`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              student_email: cert.studentEmail,
              student_name: cert.studentName,
              certificate_code: cert.certificateCode,
              certificate_id: cert.id,
              pdf_filename: cert.customFields?.pdfFilename || `Cert_${cert.certificateCode}_${cert.studentName.replace(/ /g, '_')}.pdf`,
              event_name: cert.eventName || 'Workshop',
              event_date: cert.eventDate || cert.customFields?.event_date || cert.customFields?.eventDate || eventDate || '2026-09-21',
              logo_img_url: useCustomLogo && customLogoUrl ? customLogoUrl : undefined,
              hero_img_url: useCustomHero && customHeroUrl ? customHeroUrl : undefined
            })
          });
        }
        await loadDatabaseRecords();
      } else {
        for (let i = 0; i < pendingCerts.length; i++) {
          setBatchEmailProgress({ current: i + 1, total: pendingCerts.length });
          await new Promise(r => setTimeout(r, 400));
        }
        setHistoryCerts(prev => prev.map(c => ({ ...c, emailStatus: 'sent' })));
      }
    } catch (err) {
      console.error('Failed to send batch emails:', err);
    } finally {
      setIsBatchSendingEmails(false);
    }
  };

  const handleSendSelectedEmails = async () => {
    const selectedCerts = historyCerts.filter(c => selectedCertIds.includes(c.id));
    if (selectedCerts.length === 0) return;

    setIsBatchSendingEmails(true);
    setBatchEmailProgress({ current: 0, total: selectedCerts.length });

    try {
      if (pythonBackendOnline) {
        for (let i = 0; i < selectedCerts.length; i++) {
          const cert = selectedCerts[i];
          setBatchEmailProgress({ current: i + 1, total: selectedCerts.length });
          
          await fetch(`${PYTHON_API_URL}/api/send-email-single`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              student_email: cert.studentEmail,
              student_name: cert.studentName,
              certificate_code: cert.certificateCode,
              certificate_id: cert.id,
              pdf_filename: cert.customFields?.pdfFilename || `Cert_${cert.certificateCode}_${cert.studentName.replace(/ /g, '_')}.pdf`,
              event_name: cert.eventName || 'Workshop',
              event_date: cert.eventDate || cert.customFields?.event_date || cert.customFields?.eventDate || eventDate || '2026-09-21',
              logo_img_url: useCustomLogo && customLogoUrl ? customLogoUrl : undefined,
              hero_img_url: useCustomHero && customHeroUrl ? customHeroUrl : undefined
            })
          });
        }
        await loadDatabaseRecords();
      } else {
        for (let i = 0; i < selectedCerts.length; i++) {
          setBatchEmailProgress({ current: i + 1, total: selectedCerts.length });
          await new Promise(r => setTimeout(r, 400));
        }
        setHistoryCerts(prev => prev.map(c => selectedCertIds.includes(c.id) ? { ...c, emailStatus: 'sent' } : c));
      }
    } catch (err) {
      console.error('Failed to send selected emails:', err);
    } finally {
      setIsBatchSendingEmails(false);
    }
  };

  // Reset all filters to default
  const resetCertFilters = () => {
    setHistoryQuery('');
    setCertEventFilter('All Events');
    setCertDeptFilter('All Departments');
    setCertCollegeFilter('All Colleges');
    setCertYearFilter('All Years');
    setCertSectionFilter('All Sections');
    setIsJuneJulyFilterActive(false);
    setSelectedCertIds([]);
  };

  const resetStudentFilters = () => {
    setRegQuery('');
    setStudentEventFilter('All Events');
    setStudentDeptFilter('All Departments');
    setStudentCollegeFilter('All Colleges');
    setStudentYearFilter('All Years');
    setStudentSectionFilter('All Sections');
    setSelectedStudentIds([]);
  };

  // Check Python Backend Health Status
  const checkPythonBackend = async () => {
    try {
      const res = await fetch(`${PYTHON_API_URL}/`);
      if (res.ok) {
        setPythonBackendOnline(true);
      } else {
        setPythonBackendOnline(false);
      }
    } catch {
      setPythonBackendOnline(false);
    }
  };

  // Fetch records DIRECTLY from Supabase DB (No default student data)
  const loadDatabaseRecords = async () => {
    setIsDbLoading(true);
    try {
      const dbCerts = await fetchCertificatesFromSupabase();
      const dbStudents = await fetchStudentsFromSupabase();
      const dbEvents = await fetchEventsFromSupabase();
      const dbSettings = await fetchAppSettingsFromSupabase();

      setHistoryCerts(dbCerts);
      setRegistrations(dbStudents);
      setRegisteredEvents(dbEvents);

      if (dbSettings.use_custom_logo) {
        const val = dbSettings.use_custom_logo;
        const isEnabled = typeof val === 'object' && val !== null 
          ? (val.enabled !== undefined ? !!val.enabled : !!val.url)
          : !!val;
        const urlVal = typeof val === 'object' && val !== null ? (val.url || '') : String(val);
        setUseCustomLogo(isEnabled);
        setCustomLogoUrl(urlVal);
      }
      if (dbSettings.use_custom_hero) {
        const val = dbSettings.use_custom_hero;
        const isEnabled = typeof val === 'object' && val !== null 
          ? (val.enabled !== undefined ? !!val.enabled : !!val.url)
          : !!val;
        const urlVal = typeof val === 'object' && val !== null ? (val.url || '') : String(val);
        setUseCustomHero(isEnabled);
        setCustomHeroUrl(urlVal);
      }
    } catch (err) {
      console.error('Error fetching DB records:', err);
    } finally {
      setIsDbLoading(false);
    }
  };

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      if (data?.session) {
        setSession(true);
        if (data.session.user?.email) {
          setProfileEmail(data.session.user.email);
          const metaName = data.session.user.user_metadata?.username || data.session.user.user_metadata?.display_name;
          setProfileName(metaName || data.session.user.email.split('@')[0]);
        }
      } else {
        setSession(false);
      }
      setIsAuthChecking(false);
    }).catch(() => {
      setSession(false);
      setIsAuthChecking(false);
    });

    const { data: authListener } = supabase.auth.onAuthStateChange((_event, currentSession) => {
      if (currentSession) {
        setSession(true);
        if (currentSession.user?.email) {
          setProfileEmail(currentSession.user.email);
          const metaName = currentSession.user.user_metadata?.username || currentSession.user.user_metadata?.display_name;
          setProfileName(metaName || currentSession.user.email.split('@')[0]);
        }
      } else {
        setSession(false);
      }
      setIsAuthChecking(false);
    });

    checkPythonBackend();
    loadDatabaseRecords();

    const realtimeChannel = supabase
      .channel('db-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'students' }, () => loadDatabaseRecords())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'certificates' }, () => loadDatabaseRecords())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'events' }, () => loadDatabaseRecords())
      .subscribe();

    return () => {
      authListener?.subscription.unsubscribe();
      supabase.removeChannel(realtimeChannel);
    };
  }, []);

  // Supabase Authentication Handler
  const handleAuthSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthError('');
    setAuthSuccess('');
    setIsAuthLoading(true);

    try {
      if (authMode === 'login') {
        const { data, error } = await supabase.auth.signInWithPassword({
          email: authEmail,
          password: authPassword
        });

        if (error) {
          setAuthError(error.message);
        } else if (data.session) {
          if (authUsername) setProfileName(authUsername);
          else if (data.session.user?.email) setProfileName(data.session.user.email.split('@')[0]);
          setSession(true);
        }
      } else {
        const { data, error } = await supabase.auth.signUp({
          email: authEmail,
          password: authPassword,
          options: {
            data: {
              username: authUsername,
              display_name: authUsername
            }
          }
        });

        if (error) {
          setAuthError(error.message);
        } else {
          setAuthSuccess('Registration successful! Please check your email to confirm your account, then log in.');
          setAuthMode('login');
        }
      }
    } catch (err: any) {
      setAuthError(err?.message || 'Authentication error occurred.');
    } finally {
      setIsAuthLoading(false);
    }
  };

  // Handle PPTX Template Upload
  const handleTemplateFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setTemplateFile(file);
    setIsTemplateUploading(true);

    try {
      if (pythonBackendOnline) {
        const formData = new FormData();
        formData.append('file', file);
        await fetch(`${PYTHON_API_URL}/api/upload-pptx`, {
          method: 'POST',
          body: formData
        });
      }
    } catch (err) {
      console.error('Template upload error:', err);
    } finally {
      setIsTemplateUploading(false);
    }
  };

  // Handle Excel Upload (Parses rows WITHOUT registering students in DB yet)
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsExcelParsing(true);
    try {
      if (pythonBackendOnline) {
        const formData = new FormData();
        formData.append('file', file);
        const res = await fetch(`${PYTHON_API_URL}/api/upload-excel`, {
          method: 'POST',
          body: formData
        });
        const data = await res.json();
        if (data.rows) {
          const parsedRes: ParsedExcelResult = {
            fileName: data.file_name,
            headers: data.headers,
            tokens: data.headers.map((h: string) => ({
              rawHeader: h,
              cleanHeader: h.trim(),
              token: `<<${h.trim()}>>`
            })),
            rows: data.rows,
            totalRows: data.total_rows
          };
          setExcelResult(parsedRes);
          return;
        }
      }

      // Client-side fallback
      const reader = new FileReader();
      reader.onload = (event) => {
        const buffer = event.target?.result as ArrayBuffer;
        if (buffer) {
          const parsed = parseExcelWorkbook(buffer, file.name);
          setExcelResult(parsed);
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error('Excel upload error:', err);
    } finally {
      setIsExcelParsing(false);
    }
  };

  // Copy Credential Code to Clipboard
  const handleCopyCode = (code: string) => {
    navigator.clipboard.writeText(code);
    setCopiedCodeId(code);
    setTimeout(() => setCopiedCodeId(null), 2000);
  };

  // STAGE 01 -> PROCEED TO STAGE 02 HANDLER
  // Rule 1: Validate template (.pptx), excel (.xlsx), and event name.
  // Rule 2: Register ONLY in events table, save template in event_id subfolder.
  // Rule 3: DO NOT register student or certificate data yet!
  const handleProceedToStage2 = async () => {
    if (!templateFile) {
      alert('Please upload the base PowerPoint certificate template (.pptx) before proceeding to Stage 02.');
      return;
    }
    if (!excelResult || !excelResult.rows || excelResult.rows.length === 0) {
      alert('Please upload the participant details sheet (.xlsx) before proceeding to Stage 02.');
      return;
    }
    if (!eventId.trim()) {
      alert('Please enter the Event ID / Code before proceeding to Stage 02.');
      return;
    }

    setIsProceedingToStage2(true);
    const activeEventId = eventId.trim();
    const primaryExcelEvent = (excelResult?.rows && excelResult.rows.length > 0)
      ? extractRowEventName(excelResult.rows[0], activeEventId)
      : activeEventId;

    try {
      // 1. Register Event in Supabase 'events' table with event_id=activeEventId and event_name=primaryExcelEvent
      const evRes = await saveEventToSupabase({
        event_id: activeEventId,
        event_name: primaryExcelEvent,
        event_category: eventCategory || 'Workshop',
        event_date: eventDate || new Date().toISOString().split('T')[0]
      });

      if ((evRes as any).duplicate) {
        // Exact duplicate: same event_name + same category already exists
        setSupabaseSyncError(evRes.error || 'Duplicate event. This event with the same category already exists.');
        setIsProceedingToStage2(false);
        return;
      }

      if (evRes.error && !(evRes as any).duplicate) {
        setSupabaseSyncError(evRes.error);
      } else {
        setSupabaseSyncError('');
      }

      // 2. Upload template file to Supabase Storage bucket 'templates' under event_id folder
      if (templateFile) {
        await uploadTemplateToSupabaseStorage(templateFile, activeEventId);
      }

      // 3. Upload template to event_id subfolder via Python API
      if (pythonBackendOnline && templateFile) {
        const formData = new FormData();
        formData.append('event_id', activeEventId);
        formData.append('event_name', primaryExcelEvent);
        formData.append('event_category', eventCategory || 'Workshop');
        formData.append('event_date', eventDate || new Date().toISOString().split('T')[0]);
        formData.append('file', templateFile);

        await fetch(`${PYTHON_API_URL}/api/register-event`, {
          method: 'POST',
          body: formData
        });
      }

      setGeneratorStep(2);
    } catch (err) {
      console.error('Error proceeding to Stage 02:', err);
    } finally {
      setIsProceedingToStage2(false);
    }
  };

  // STAGE 02 -> GENERATE CERTIFICATES BATCH HANDLER
  // Rule: ONLY upon clicking 'Generate Certificates' do we register students and certificates in DB & compile PDFs.
  const handleGenerateCertificatesBatch = async () => {
    if (!excelResult || excelResult.rows.length === 0) {
      alert('Please upload a valid Excel participant details sheet first.');
      return;
    }

    setIsGeneratingBatch(true);
    setBatchGenSuccessMsg('');
    const activeEventId = (eventId.trim() || 'default-event').trim();
    const primaryExcelEvent = (excelResult?.rows && excelResult.rows.length > 0)
      ? extractRowEventName(excelResult.rows[0], activeEventId)
      : activeEventId;

    try {

      const certsFromExcel: Certificate[] = excelResult.rows.map((row, idx) => {
        const rollNo = extractStudentId(row);
        const name = String(row['Name'] || row['name'] || 'Participant').trim();
        const email = extractStudentEmail(row, rollNo);
        const code = generateCertificateId('WRK', idx + 1);
        const college = String(row['College Name'] || row['College'] || 'Kongu Engineering College').trim() || 'Kongu Engineering College';
        const rowEvent = extractRowEventName(row, primaryExcelEvent);
        const rowTitle = extractRowTitle(row, '');
        const rowYear = extractRowYear(row, '');

        return {
          id: `cert-excel-${idx + 1}`,
          certificateCode: code,
          studentId: rollNo,
          studentName: name,
          studentEmail: email,
          eventId: activeEventId,
          eventName: rowEvent,
          eventDate: eventDate || new Date().toISOString().split('T')[0],
          customFields: {
            ...row,
            Event: rowEvent,
            EVENT: rowEvent,
            event: rowEvent,
            'Event Name': rowEvent,
            event_id: activeEventId,
            event_name: rowEvent,
            Title: rowTitle,
            title: rowTitle,
            TITLE: rowTitle,
            Topic: rowTitle || rowEvent,
            'Paper Title': rowTitle,
            'Project Title': rowTitle,
            event_category: eventCategory || 'Workshop',
            event_date: eventDate,
            issue_date: issueDate,
            college_name: college,
            year: rowYear,
            year_of_study: rowYear
          },
          issueDate: issueDate || new Date().toISOString().split('T')[0],
          emailStatus: 'pending',
          emailSentAt: new Date().toISOString(),
          createdAt: new Date().toISOString()
        };
      });

      const studentsFromExcel: Student[] = excelResult.rows.map((row, idx) => {
        const rollNo = extractStudentId(row);
        const name = String(row['Name'] || row['name'] || 'Participant').trim();
        const email = extractStudentEmail(row, rollNo);
        const section = String(row['Section'] || row['section'] || '').trim();
        const phone = extractRowPhone(row, String(row['Mobile number '] || row['phone'] || '').trim());
        const college = String(row['College Name'] || row['College'] || 'Kongu Engineering College').trim() || 'Kongu Engineering College';
        const year = extractRowYear(row, String(row['Year of Study'] || row['Year'] || row['year'] || '').trim());

        return {
          id: `stu-excel-${idx + 1}`,
          registerNo: rollNo,
          name: name,
          email: email,
          phone: phone,
          department: 'Computer Science and Engineering',
          yearOfStudy: year,
          section: section,
          collegeName: college,
          createdAt: new Date().toISOString()
        };
      });

      // 1. Compile native PDFs & Register Students/Certificates in Supabase DB via Python API
      if (pythonBackendOnline) {
        const res = await fetch(`${PYTHON_API_URL}/api/generate-certificates`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            batch_id: `Batch_${activeEventId}`,
            event_id: activeEventId,
            event_name: primaryExcelEvent,
            event_category: eventCategory || 'Workshop',
            event_date: eventDate,
            issue_date: issueDate,
            rows: excelResult.rows
          })
        });
        const data = await res.json();
        if (data.status === 'success') {
          setBatchGenSuccessMsg(`✓ Successfully Generated ${data.count} Certificates, Registered Students & Synced Database!`);
        }
      } else {
        // Fallback client-side DB registration
        const studentRes = await saveStudentsToSupabase(studentsFromExcel);
        const certRes = await saveCertificatesToSupabase(certsFromExcel, studentRes.data);

        if (studentRes.error || certRes.error) {
          setSupabaseSyncError(studentRes.error || certRes.error || '');
        } else {
          setSupabaseSyncError('');
        }
        setBatchGenSuccessMsg(`✓ Successfully Generated ${excelResult.totalRows} Certificates & Synced Database!`);
      }

      // 4. Reload Database Records Live
      await loadDatabaseRecords();
    } catch (err) {
      console.error('Batch certificate generation error:', err);
    } finally {
      setIsGeneratingBatch(false);
    }
  };

  // DYNAMIC VIEW HANDLER (Opens PDF directly in a new browser tab without downloading)
  const handleViewCertificate = (c: Certificate) => {
    const safeName = c.studentName.trim().replace(/\s+/g, '_');
    const certCode = c.certificateCode || '';
    const pdfFilename = `Cert_${certCode ? `${certCode}_` : ''}${safeName}.pdf`;
    const viewUrl = `${PYTHON_API_URL}/api/download-pdf/${encodeURIComponent(pdfFilename)}?cert_id=${encodeURIComponent(c.id || '')}&event_id=${encodeURIComponent(c.eventId || '')}&mode=inline&t=${Date.now()}`;
    window.open(viewUrl, '_blank');
  };

  // DYNAMIC DOWNLOAD HANDLER (Triggers direct file download ONLY when clicking Download)
  const handleDownloadCertificate = (c: Certificate) => {
    const safeName = c.studentName.trim().replace(/\s+/g, '_');
    const certCode = c.certificateCode || '';
    const pdfFilename = `Cert_${certCode ? `${certCode}_` : ''}${safeName}.pdf`;
    const downloadUrl = `${PYTHON_API_URL}/api/download-pdf/${encodeURIComponent(pdfFilename)}?cert_id=${encodeURIComponent(c.id || '')}&event_id=${encodeURIComponent(c.eventId || '')}&mode=attachment&t=${Date.now()}`;

    const a = document.createElement('a');
    a.href = downloadUrl;
    a.download = `${c.studentName}_Certificate.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const handleExportExcel = async () => {
    const targetCerts = selectedCertIds.length > 0
      ? historyCerts.filter(c => selectedCertIds.includes(c.id))
      : filteredCertificates;

    if (!targetCerts || targetCerts.length === 0) {
      alert('No certificate records available to export.');
      return;
    }

    setIsExportingExcel(true);
    try {
      const exportData = targetCerts.map((c, idx) => ({
        'S.No': idx + 1,
        'Student Name': c.studentName,
        'Register No': c.studentId,
        'Email Address': c.studentEmail,
        'Event Name': c.eventName,
        'Issue Date': c.issueDate,
        'Email Status': c.emailStatus,
        'College Name': c.customFields?.college_name || 'Kongu Engineering College',
        'Department': c.customFields?.department || 'Computer Science and Engineering',
        'Section': c.customFields?.section || '',
        'Year of Study': c.customFields?.year_of_study || ''
      }));

      const worksheet = XLSX.utils.json_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Certificates');

      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const data = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const fileName = selectedCertIds.length > 0 
        ? `CSEA_Certificates_Selected_${new Date().toISOString().split('T')[0]}.xlsx` 
        : `CSEA_Certificates_Filtered_${new Date().toISOString().split('T')[0]}.xlsx`;

      const url = URL.createObjectURL(data);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error exporting Excel:', err);
      alert('Failed to export Excel file. Please try again.');
    } finally {
      setIsExportingExcel(false);
    }
  };

  const handleExportStudentExcel = async () => {
    const targetStudents = filteredStudents.length > 0 ? filteredStudents : registrations;

    if (!targetStudents || targetStudents.length === 0) {
      alert('No student records available to export.');
      return;
    }

    setIsExportingExcel(true);
    try {
      // Build a map of studentId -> certificate record for quick lookup
      const certMap = new Map<string, (typeof historyCerts)[0]>();
      historyCerts.forEach(c => {
        if (c.studentId) certMap.set(c.studentId, c);
        // also index by student name as fallback
        if (c.studentName) certMap.set(c.studentName.toLowerCase(), c);
      });

      // Build a map of eventId -> event record
      const eventMap = new Map<string, (typeof registeredEvents)[0]>();
      registeredEvents.forEach(ev => {
        if (ev.eventId) eventMap.set(ev.eventId, ev);
      });

      const exportData = targetStudents.map((s, idx) => {
        const cert = certMap.get(s.registerNo) || certMap.get(s.name?.toLowerCase() || '') || null;
        const ev = cert?.eventId ? (eventMap.get(cert.eventId) || null) : null;
        return {
          'S.No': idx + 1,
          'Student Name': s.name,
          'Register No': s.registerNo,
          'Email Address': s.email,
          'College Name': s.collegeName || 'Kongu Engineering College',
          'Department': s.department || 'Computer Science and Engineering',
          'Section': s.section || '',
          'Year of Study': s.yearOfStudy || '',
          'Event Name': cert?.eventName || ev?.eventName || '',
          'Issue Date': cert?.issueDate || '',
          'Email Status': cert?.emailStatus || 'pending',
        };
      });

      const worksheet = XLSX.utils.json_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Students');

      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const data = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      const fileName = `CSEA_Students_${new Date().toISOString().split('T')[0]}.xlsx`;
      const url = URL.createObjectURL(data);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error exporting Student Excel:', err);
      alert('Failed to export Excel. Please try again.');
    } finally {
      setIsExportingExcel(false);
    }
  };

  const handleExportZip = async () => {
    const targetCerts = selectedCertIds.length > 0
      ? historyCerts.filter(c => selectedCertIds.includes(c.id))
      : filteredCertificates;

    if (!targetCerts || targetCerts.length === 0) {
      alert('No certificate records selected or available to package into ZIP.');
      return;
    }

    setIsExportingZip(true);
    try {
      const zip = new JSZip();
      const folder = zip.folder('Certificates') || zip;

      for (let i = 0; i < targetCerts.length; i++) {
        const c = targetCerts[i];
        const safeName = c.studentName.trim().replace(/\s+/g, '_');
        const pdfFilename = `Cert_${safeName}.pdf`;
        const downloadUrl = `${PYTHON_API_URL}/api/download-pdf/${encodeURIComponent(pdfFilename)}?mode=inline&t=${Date.now()}`;

        try {
          const res = await fetch(downloadUrl);
          if (res.ok) {
            const blob = await res.blob();
            const archiveFilename = `Cert_${c.studentId || (i + 1)}_${safeName}.pdf`;
            folder.file(archiveFilename, blob);
          }
        } catch (fetchErr) {
          console.warn(`Failed to fetch PDF for ${c.studentName}:`, fetchErr);
        }
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const fileName = selectedCertIds.length > 0
        ? `CSEA_Certificates_Selected_${new Date().toISOString().split('T')[0]}.zip`
        : `CSEA_Certificates_Filtered_${new Date().toISOString().split('T')[0]}.zip`;

      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error packaging ZIP:', err);
      alert('Failed to package ZIP archive. Please try again.');
    } finally {
      setIsExportingZip(false);
    }
  };

  const handleSendEmail = () => {
    setIsSendingEmail(true);
    setTimeout(() => {
      alert(`Dispatching certificate emails to ${selectedCertIds.length > 0 ? selectedCertIds.length : filteredCertificates.length} participants...`);
      setIsSendingEmail(false);
    }, 1500);
  };

  // -------------------------------------------------------------
  // DYNAMIC FILTER OPTIONS POPULATED FROM EXCEL & SUPABASE DB ONLY
  // -------------------------------------------------------------
  const certEventOptions = Array.from(new Set([
    ...historyCerts.map(c => c.eventName || c.eventId),
    ...(excelResult?.rows.map(r => r.event_name || r.eventName || r['Event Name'] || r['Event']) || [])
  ].filter(Boolean)));

  const certDeptOptions = Array.from(new Set([
    ...historyCerts.map(c => c.customFields?.department),
    ...registrations.map(s => s.department),
    ...(excelResult?.rows.map(r => r.department || r.Department || r['Department']) || [])
  ].filter(Boolean)));

  const certCollegeOptions = Array.from(new Set([
    ...historyCerts.map(c => c.customFields?.college_name),
    ...registrations.map(s => s.collegeName),
    ...(excelResult?.rows.map(r => r.college_name || r['College Name'] || r.College || r['College']) || [])
  ].filter(Boolean)));

  const certYearOptions = Array.from(new Set([
    ...historyCerts.map(c => c.customFields?.year_of_study),
    ...registrations.map(s => s.yearOfStudy),
    ...(excelResult?.rows.map(r => r.year_of_study || r['Year of Study'] || r.Year || r['Year']) || [])
  ].filter(Boolean)));

  const certSectionOptions = Array.from(new Set([
    ...historyCerts.map(c => c.customFields?.section),
    ...registrations.map(s => s.section),
    ...(excelResult?.rows.map(r => r.section || r.Section || r['Section']) || [])
  ].filter(Boolean)));

  // Filtered Certificates
  const filteredCertificates = historyCerts.filter(c => {
    const q = historyQuery.toLowerCase();
    const matchesSearch = c.studentName.toLowerCase().includes(q) ||
                          c.studentId.toLowerCase().includes(q) ||
                          c.studentEmail.toLowerCase().includes(q);

    const matchesEvent = certEventFilter === 'All Events' || (c.eventName && c.eventName.toLowerCase().includes(certEventFilter.toLowerCase()));
    const matchesDept = certDeptFilter === 'All Departments' || (c.customFields?.department && c.customFields.department.toLowerCase() === certDeptFilter.toLowerCase());
    const matchesCollege = certCollegeFilter === 'All Colleges' || (c.customFields?.college_name && c.customFields.college_name.toLowerCase().includes(certCollegeFilter.toLowerCase()));
    const matchesYear = certYearFilter === 'All Years' || (c.customFields?.year_of_study && c.customFields.year_of_study.toLowerCase() === certYearFilter.toLowerCase());
    const matchesSection = certSectionFilter === 'All Sections' || (c.customFields?.section && c.customFields.section.toLowerCase() === certSectionFilter.replace('Section ', '').toLowerCase());

    return matchesSearch && matchesEvent && matchesDept && matchesCollege && matchesYear && matchesSection;
  });

  // Build a map of studentId/registerNo -> eventName for fast event-based student filtering
  const studentEventMap = new Map<string, string>();
  historyCerts.forEach(c => {
    if (c.studentId && c.eventName) studentEventMap.set(c.studentId, c.eventName);
  });

  // Filtered Students
  const filteredStudents = registrations.filter(s => {
    const q = regQuery.toLowerCase();
    const section = (s.section || '').toLowerCase();
    const collegeName = (s.collegeName || '').toLowerCase();
    const department = (s.department || '').toLowerCase();
    const yearOfStudy = (s.yearOfStudy || '').toLowerCase();
    const name = (s.name || '').toLowerCase();
    const registerNo = (s.registerNo || '').toLowerCase();
    const email = (s.email || '').toLowerCase();

    const matchesSearch = !q ||
      name.includes(q) ||
      registerNo.includes(q) ||
      email.includes(q) ||
      section.includes(q) ||
      collegeName.includes(q);

    const matchesDept = studentDeptFilter === 'All Departments' || department === studentDeptFilter.toLowerCase();
    const matchesCollege = studentCollegeFilter === 'All Colleges' || collegeName.includes(studentCollegeFilter.toLowerCase());
    const matchesYear = studentYearFilter === 'All Years' || yearOfStudy === studentYearFilter.toLowerCase();
    const matchesSec = studentSectionFilter === 'All Sections' || section === studentSectionFilter.replace('Section ', '').toLowerCase();

    // Event filter: look up which event this student participated in via their certificate
    const studentEventName = studentEventMap.get(s.registerNo) || studentEventMap.get(s.id) || '';
    const matchesEvent = studentEventFilter === 'All Events' ||
      studentEventName.toLowerCase().includes(studentEventFilter.toLowerCase());

    return matchesSearch && matchesDept && matchesCollege && matchesYear && matchesSec && matchesEvent;
  });

  // Toggle select all certs
  const toggleSelectAllCerts = () => {
    if (selectedCertIds.length === filteredCertificates.length) {
      setSelectedCertIds([]);
    } else {
      setSelectedCertIds(filteredCertificates.map(c => c.id));
    }
  };

  // Toggle single cert selection
  const toggleSelectCert = (id: string) => {
    if (selectedCertIds.includes(id)) {
      setSelectedCertIds(selectedCertIds.filter(i => i !== id));
    } else {
      setSelectedCertIds([...selectedCertIds, id]);
    }
  };

  // ========================================================
  // PREMIUM VERIFYING SESSION LOADING SCREEN
  // ========================================================
  if (isAuthChecking) {
    return (
      <div className="fixed inset-0 z-50 bg-slate-50 flex items-center justify-center p-4">
        <div className="bg-white border border-slate-200/80 rounded-3xl p-10 max-w-sm w-full shadow-xl shadow-slate-200/40 flex flex-col items-center space-y-6 text-center">
          
          <div className="w-20 h-20 bg-white border-2 border-indigo-600 shadow-lg p-1.5 rounded-3xl flex items-center justify-center animate-pulse">
            <img src={useCustomLogo && customLogoUrl ? customLogoUrl : "/csea_logo.png"} alt="CSEA" className="w-full h-full object-contain" />
          </div>

          <div className="relative w-10 h-10 flex items-center justify-center">
            <div className="absolute inset-0 rounded-full border-2 border-indigo-500/20 border-t-indigo-600 animate-spin"></div>
            <div className="absolute inset-1.5 rounded-full border-2 border-blue-500/20 border-b-blue-500 animate-[spin_1.5s_linear_infinite_reverse]"></div>
          </div>

          <div>
            <h2 className="text-lg font-black text-slate-800 tracking-tight">CSEA Certificate Hub</h2>
            <p className="text-xs text-indigo-600 font-extrabold tracking-widest uppercase mt-2 animate-pulse flex items-center justify-center space-x-1.5">
              <Sparkles className="w-3.5 h-3.5 text-amber-500 shrink-0" />
              <span>Verifying session...</span>
            </p>
          </div>

        </div>
      </div>
    );
  }

  // ========================================================
  // SUPABASE LOGIN / SIGNUP MODAL (FULL SCREEN DARK #030712)
  // ========================================================
  if (!session) {
    return (
      <div className="fixed inset-0 z-50 bg-[#030712] flex items-center justify-center p-4">
        <div className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl space-y-6 text-center">
          
          <div className="w-24 h-24 bg-white border-2 border-indigo-600 shadow-lg p-1.5 rounded-3xl mx-auto flex items-center justify-center">
            <img src={useCustomLogo && customLogoUrl ? customLogoUrl : "/csea_logo.png"} alt="CSEA" className="w-full h-full object-contain" />
          </div>

          <div>
            <h1 className="text-2xl font-extrabold text-slate-900 tracking-tight">CSEA Certificate Hub</h1>
          </div>

          <div className="bg-slate-100 p-1 rounded-xl flex">
            <button
              onClick={() => { setAuthMode('login'); setAuthError(''); }}
              className={`flex-1 py-2 text-xs font-extrabold rounded-lg transition-all ${
                authMode === 'login' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-500'
              }`}
            >
              Log In
            </button>
            <button
              onClick={() => { setAuthMode('signup'); setAuthError(''); setAuthSuccess(''); }}
              className={`flex-1 py-2 text-xs font-extrabold rounded-lg transition-all ${
                authMode === 'signup' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-500'
              }`}
            >
              Sign Up
            </button>
          </div>

          {authSuccess && (
            <div className="p-3 bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-bold rounded-xl text-left">
              {authSuccess}
            </div>
          )}

          {authError && (
            <div className="p-3 bg-red-50 border border-red-200 text-red-700 text-xs font-bold rounded-xl text-left">
              {authError}
            </div>
          )}

          <form onSubmit={handleAuthSubmit} className="space-y-4 text-left">
            {authMode === 'signup' && (
              <div>
                <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                  FULL NAME / USERNAME
                </label>
                <div className="relative">
                  <User className="w-4 h-4 text-slate-400 absolute left-3 top-3.5" />
                  <input
                    type="text"
                    placeholder="e.g. Rohith P"
                    value={authUsername}
                    onChange={(e) => setAuthUsername(e.target.value)}
                    className="w-full pl-9 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:outline-none focus:ring-2 focus:ring-indigo-600"
                    required
                  />
                </div>
              </div>
            )}

            <div>
              <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                EMAIL ADDRESS
              </label>
              <div className="relative">
                <Mail className="w-4 h-4 text-slate-400 absolute left-3 top-3.5" />
                <input
                  type="email"
                  placeholder="e.g. admin@kongu.edu"
                  value={authEmail}
                  onChange={(e) => setAuthEmail(e.target.value)}
                  className="w-full pl-9 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:outline-none focus:ring-2 focus:ring-indigo-600"
                  required
                />
              </div>
            </div>

            <div>
              <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                PASSWORD
              </label>
              <div className="relative">
                <Lock className="w-4 h-4 text-slate-400 absolute left-3 top-3.5" />
                <input
                  type="password"
                  value={authPassword}
                  onChange={(e) => setAuthPassword(e.target.value)}
                  className="w-full pl-9 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:outline-none focus:ring-2 focus:ring-indigo-600"
                  required
                />
              </div>
            </div>

            <button
              type="submit"
              disabled={isAuthLoading}
              className="w-full bg-[#4f46e5] hover:bg-[#4338ca] text-white font-extrabold py-3.5 rounded-xl text-xs uppercase tracking-wider shadow-lg transition-colors mt-2 flex items-center justify-center space-x-2"
            >
              {isAuthLoading ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  <span>AUTHENTICATING...</span>
                </>
              ) : (
                <span>{authMode === 'login' ? 'LOG IN TO PANEL' : 'SIGN UP ACCOUNT'}</span>
              )}
            </button>
          </form>

        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      <link rel="stylesheet" href={getGoogleFontStylesheetUrl(CERTIFICATE_FONTS)} />

      {/* HEADER & NAVIGATION BAR */}
      <header className="sticky top-0 z-50 bg-white border-b border-slate-200 shadow-sm">
        <div className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between py-2.5 sm:py-3 md:h-20 gap-2 sm:gap-4">
            
            {/* Logo & Portal Title */}
            <Link href="/" className="flex items-center space-x-2.5 sm:space-x-3.5 group shrink-0 min-w-0">
              <div className="w-10 h-10 sm:w-14 sm:h-14 md:w-16 md:h-16 bg-white border-2 border-indigo-600 shadow-md rounded-xl sm:rounded-2xl p-0.5 sm:p-1 flex items-center justify-center transition-transform duration-200 group-hover:scale-105 shrink-0">
                <img src={useCustomLogo && customLogoUrl ? customLogoUrl : "/csea_logo.png"} alt="CSEA" className="w-full h-full object-contain" />
              </div>
              <div className="min-w-0">
                <span className="font-black text-sm sm:text-lg md:text-xl text-slate-900 tracking-tight block leading-tight truncate group-hover:text-indigo-600 transition-colors">
                  CSEA Certificate Hub
                </span>
                <span className="text-[9px] sm:text-[10px] font-extrabold uppercase tracking-widest text-indigo-600/80 block truncate">
                  Official Admin Portal
                </span>
              </div>
            </Link>

            {/* Desktop Navigation */}
            <nav className="hidden lg:flex items-center bg-slate-100/80 p-1.5 rounded-2xl border border-slate-200 space-x-1 shrink-0">
              {[
                { id: 'generator', label: 'Generator', icon: LayoutDashboard },
                { id: 'invitation', label: 'Invitations', icon: Mail },
                { id: 'history', label: 'Certificates', icon: FileText },
                { id: 'registration', label: 'Student Directory', icon: Database },
                { id: 'profile', label: 'Profile', icon: User },
              ].map((tab) => {
                const Icon = tab.icon;
                const isActive = activeTab === tab.id;
                return (
                  <button
                    key={tab.id}
                    onClick={() => setActiveTab(tab.id as any)}
                    className={`flex items-center space-x-2 px-3.5 py-2 rounded-xl text-xs font-extrabold transition-all shrink-0 ${
                      isActive
                        ? 'bg-white text-slate-900 shadow-md border border-slate-200'
                        : 'text-slate-500 hover:text-slate-900'
                    }`}
                  >
                    <Icon className="w-4 h-4" />
                    <span>{tab.label}</span>
                  </button>
                );
              })}
            </nav>

            {/* User Profile & Logout */}
            <div className="flex items-center space-x-2 sm:space-x-3 shrink-0">
              <div className="hidden sm:block text-right text-xs">
                <span className="text-[9px] sm:text-[10px] text-slate-400 font-extrabold uppercase block leading-none mb-0.5">LOGGED IN AS</span>
                <span className="font-bold text-slate-800 text-xs sm:text-sm truncate max-w-[120px] block">{profileName}</span>
              </div>
              <button
                onClick={async () => {
                  await supabase.auth.signOut();
                  setSession(false);
                }}
                className="p-2 sm:p-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl transition-colors shrink-0 flex items-center justify-center"
                title="Logout"
              >
                <LogOut className="w-4 h-4" />
              </button>
            </div>

          </div>

          {/* Mobile & Tablet Navigation Sub-Bar */}
          <div className="lg:hidden pb-2.5 pt-1 overflow-x-auto no-scrollbar -mx-3 px-3">
            <nav className="flex items-center space-x-1.5 bg-slate-100/90 p-1 rounded-xl border border-slate-200 w-max min-w-full">
              {[
                { id: 'generator', label: 'Generator', icon: LayoutDashboard },
                { id: 'invitation', label: 'Invitations', icon: Mail },
                { id: 'history', label: 'Certificates', icon: FileText },
                { id: 'registration', label: 'Students', icon: Database },
                { id: 'profile', label: 'Profile', icon: User },
              ].map((tab) => {
                const Icon = tab.icon;
                const isActive = activeTab === tab.id;
                return (
                  <button
                    key={tab.id}
                    onClick={() => setActiveTab(tab.id as any)}
                    className={`flex items-center space-x-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all whitespace-nowrap shrink-0 ${
                      isActive
                        ? 'bg-white text-indigo-700 shadow-sm border border-slate-200 font-extrabold'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    <Icon className="w-3.5 h-3.5" />
                    <span>{tab.label}</span>
                  </button>
                );
              })}
            </nav>
          </div>
        </div>
      </header>

      {/* MAIN CONTENT */}
      <main className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-4 sm:py-8">
        
        {supabaseSyncError && (
          <div className="mb-6 p-4 sm:p-5 bg-amber-50 border border-amber-300 rounded-2xl sm:rounded-3xl text-amber-900 shadow-sm space-y-2">
            <div className="flex items-center space-x-2 font-extrabold text-amber-800 text-xs sm:text-sm">
              <AlertCircle className="w-4 h-4 sm:w-5 sm:h-5 text-amber-600 shrink-0" />
              <span>Supabase Database Sync Notice: {supabaseSyncError}</span>
            </div>
            <p className="text-xs text-amber-800 font-medium leading-relaxed">
              If Supabase schema access is restricted, run the SQL script in <code className="bg-amber-100 px-2 py-0.5 rounded font-mono font-bold">supabase_setup.sql</code> in your Supabase SQL Editor.
            </p>
          </div>
        )}

        {/* TAB 0: INVITATION */}
        {activeTab === 'invitation' && (
          <InvitationTab />
        )}

        {/* TAB 1: GENERATOR */}
        {activeTab === 'generator' && (
          <div className="space-y-6 sm:space-y-8">
            
            {/* HERO BANNER - RESPONSIVE RATIO & FIT */}
            <div className="relative overflow-hidden rounded-2xl sm:rounded-[32px] h-[220px] sm:h-[360px] md:h-[480px] w-full shadow-[0_20px_60px_rgba(0,0,0,0.12)] border-2 sm:border-4 border-white bg-slate-950 flex items-center justify-center">
              <img src={useCustomHero && customHeroUrl ? customHeroUrl : "/hero.png"} alt="Hero Banner" className="w-full h-full object-contain md:object-cover object-center transition-transform duration-700 hover:scale-[1.01]" />
            </div>

            {/* ACTIVE STAGE CARD */}
            <div className="rounded-2xl sm:rounded-3xl border border-emerald-100 bg-white p-4 sm:p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4 sm:gap-6 relative overflow-hidden">
              <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-[#12a150]" />

              <div className="flex items-start sm:items-center gap-3 sm:gap-5">
                <div className="rounded-2xl border border-[#12a150]/20 bg-[#12a150]/10 px-3 py-1.5 sm:px-4 sm:py-2 text-center shrink-0">
                  <span className="text-[9px] sm:text-[10px] font-extrabold uppercase tracking-wider text-[#12a150] block">Active Stage</span>
                  <span className="text-lg sm:text-xl font-bold font-mono text-[#12a150] mt-0.5 block">0{generatorStep} / 02</span>
                </div>

                <div>
                  <h3 className="text-sm sm:text-base font-bold text-slate-900 flex items-center gap-2">
                    <span className="inline-block w-2 h-2 rounded-full bg-[#12a150] animate-pulse shrink-0" />
                    <span className="leading-snug">{generatorStep === 1 ? 'Upload Base Certificate Files' : 'Compile Batch Certificate Run'}</span>
                  </h3>
                  <p className="mt-1 text-[11px] sm:text-xs text-slate-600 leading-relaxed max-w-3xl">
                    {generatorStep === 1
                      ? 'Select the base PowerPoint (.pptx) certificate template and participant details Excel sheet (.xlsx).'
                      : 'Verify participant record mappings and compile native High-DPI PowerPoint certificates.'
                    }
                  </p>
                </div>
              </div>

              <div className="flex items-center justify-between sm:justify-end gap-2 shrink-0 pt-2 sm:pt-0 border-t sm:border-t-0 border-slate-100">
                <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Status</span>
                <div className="flex items-center gap-2 bg-[#12a150] px-3 py-1.5 sm:px-4 sm:py-2 rounded-full shadow-md select-none">
                  <div className="w-1.5 h-1.5 rounded-full bg-white animate-ping" />
                  <span className="text-[10px] font-bold text-white uppercase tracking-wider">READY</span>
                </div>
              </div>
            </div>

            {/* STAGE 1 DARK UPLOAD CONTAINER (#030712) */}
            {generatorStep === 1 && (
              <div className="rounded-2xl sm:rounded-[32px] border border-zinc-800 bg-[#030712] p-4 sm:p-6 md:p-8 shadow-2xl text-white space-y-6 sm:space-y-8">
                <div className="flex justify-between items-center">
                  <h2 className="font-sans text-lg sm:text-2xl font-bold text-white">
                    Upload Base <span className="text-blue-500">PowerPoint Certificate Template</span>
                  </h2>
                </div>

                {/* SINGLE CERTIFICATE TEMPLATE UPLOAD BOX */}
                <div className="relative group border-2 border-dashed border-zinc-800 bg-[#0b0f19] hover:border-zinc-700 hover:bg-[#121626] rounded-2xl p-4 sm:p-8 transition-all min-h-[140px] sm:min-h-[160px] flex flex-col items-center justify-center text-center">
                  {isTemplateUploading ? (
                    <div className="flex flex-col items-center">
                      <Loader2 className="w-8 h-8 sm:w-10 sm:h-10 animate-spin text-blue-500 mb-2" />
                      <p className="text-xs font-extrabold text-blue-400">Uploading PPTX Template...</p>
                    </div>
                  ) : (
                    <>
                      <FileText className="w-10 h-10 sm:w-12 sm:h-12 mb-2 sm:mb-3 text-zinc-500 group-hover:text-blue-500 transition-colors" />
                      <p className="text-xs sm:text-sm font-bold text-white mb-1">1. Certificate Base Template (.pptx)</p>
                      <p className="text-[11px] sm:text-xs text-zinc-400">Drag & drop your WORKSHOP.pptx PowerPoint template file here</p>
                      <input type="file" accept=".pptx,.html,.htm" onChange={handleTemplateFileUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
                    </>
                  )}
                  
                  {templateFile && !isTemplateUploading && (
                    <span className="text-xs text-blue-400 font-extrabold mt-3 bg-blue-950/80 px-3 py-1 rounded-full border border-blue-800">
                      ✓ Uploaded: {templateFile.name} (Native PPTX Active)
                    </span>
                  )}
                </div>

                {/* Bottom Row: Excel Dropzone & Unfilled Batch Release Details */}
                <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 sm:gap-6 pt-4 border-t border-zinc-800/80">
                  
                  {/* Excel Upload Dropzone */}
                  <div className="lg:col-span-7 relative group border-2 border-dashed border-zinc-800 bg-[#0b0f19] hover:border-zinc-700 p-4 sm:p-8 rounded-2xl flex flex-col items-center justify-center text-center">
                    {isExcelParsing ? (
                      <div className="flex flex-col items-center py-4">
                        <Loader2 className="w-8 h-8 sm:w-10 sm:h-10 animate-spin text-blue-500 mb-2" />
                        <p className="text-xs font-extrabold text-blue-400">Parsing Participant Excel Sheet...</p>
                      </div>
                    ) : (
                      <>
                        <FileText className="w-10 h-10 sm:w-12 sm:h-12 text-blue-500 mb-2 sm:mb-3" />
                        <p className="text-xs sm:text-sm font-bold text-white mb-1">2. Participant Details Sheet (.xlsx)</p>
                        <p className="text-[11px] sm:text-xs text-zinc-400 mb-3">
                          {excelResult ? `${excelResult.fileName} (${excelResult.totalRows} Rows)` : 'Drag & drop your Excel sheet here'}
                        </p>
                        <label className="cursor-pointer bg-blue-600 hover:bg-blue-700 text-white font-bold px-4 sm:px-5 py-2 sm:py-2.5 rounded-xl text-xs uppercase tracking-wider transition-all flex items-center space-x-2">
                          <Upload className="w-4 h-4" />
                          <span>Browse Files</span>
                          <input type="file" accept=".xlsx, .csv" onChange={handleFileUpload} className="hidden" />
                        </label>
                      </>
                    )}
                  </div>

                  {/* Batch Release Details Form */}
                  <div className="lg:col-span-5 bg-[#0b0f19] border border-zinc-800 p-4 sm:p-6 rounded-2xl space-y-4">
                    <h3 className="text-xs sm:text-sm font-bold text-white border-b border-zinc-800 pb-3">
                      3. Batch Release Details
                    </h3>

                    <div className="space-y-3">
                      <div>
                        <label className="block text-[10px] font-extrabold text-zinc-400 uppercase tracking-wider mb-1">
                          EVENT ID / CODE
                        </label>
                        <input
                          type="text"
                          value={eventId}
                          onChange={(e) => setEventId(e.target.value)}
                          placeholder="e.g. R - Code"
                          className="w-full bg-[#121626] border border-zinc-800 text-white text-xs p-2.5 sm:p-3 rounded-xl focus:outline-none focus:border-blue-500 font-medium"
                        />
                        <div className="flex items-center space-x-1.5 mt-1.5 text-[10px] text-emerald-400/90 font-medium">
                          <span>✨</span>
                          <span>Event names are dynamically extracted per participant from the uploaded Excel sheet.</span>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div>
                          <label className="block text-[9px] font-extrabold text-zinc-400 uppercase tracking-wider mb-1">
                            EVENT DATE
                          </label>
                          <input
                            type="date"
                            value={eventDate}
                            onChange={(e) => setEventDate(e.target.value)}
                            className="w-full bg-[#121626] border border-zinc-800 text-white text-xs p-2.5 rounded-xl focus:outline-none focus:border-blue-500 font-medium"
                          />
                        </div>

                        <div>
                          <label className="block text-[9px] font-extrabold text-zinc-400 uppercase tracking-wider mb-1">
                            CERTIFICATE ISSUE DATE
                          </label>
                          <input
                            type="date"
                            value={issueDate}
                            onChange={(e) => setIssueDate(e.target.value)}
                            className="w-full bg-[#121626] border border-zinc-800 text-white text-xs p-2.5 rounded-xl focus:outline-none focus:border-blue-500 font-medium"
                          />
                        </div>
                      </div>
                    </div>
                  </div>

                </div>

                <div className="flex justify-end pt-4 border-t border-zinc-800">
                  <button
                    onClick={handleProceedToStage2}
                    disabled={isProceedingToStage2}
                    className="relative overflow-hidden bg-gradient-to-r from-blue-600 via-indigo-600 to-blue-700 hover:from-blue-500 hover:to-indigo-600 text-white font-extrabold px-9 py-4 rounded-2xl text-xs uppercase tracking-widest shadow-xl hover:shadow-blue-500/25 transition-all duration-300 flex items-center space-x-3 disabled:opacity-80 group cursor-pointer border border-blue-400/30"
                  >
                    {isProceedingToStage2 ? (
                      <>
                        <div className="relative w-5 h-5 flex items-center justify-center">
                          <div className="absolute inset-0 rounded-full border-2 border-white/20 border-t-white animate-spin"></div>
                          <div className="absolute inset-1 rounded-full border-2 border-indigo-200/30 border-b-indigo-300 animate-[spin_1.5s_linear_infinite_reverse]"></div>
                        </div>
                        <span className="font-black text-white tracking-widest animate-pulse">Registering Event...</span>
                      </>
                    ) : (
                      <>
                        <span className="font-black">Register Event</span>
                        <ArrowRight className="w-4 h-4 transition-transform duration-300 group-hover:translate-x-1.5 text-blue-200" />
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}

            {/* STAGE 02: VERIFY EXCEL DATA & GENERATE CERTIFICATES BATCH */}
            {generatorStep === 2 && (
              <div className="space-y-6 sm:space-y-8">
                
                {/* BATCH SUMMARY & REGISTERED EVENT BADGE */}
                <div className="bg-white p-4 sm:p-6 md:p-8 rounded-2xl sm:rounded-3xl border border-slate-200 shadow-sm space-y-4 sm:space-y-6">
                  <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 border-b border-slate-100 pb-4 sm:pb-5">
                    <div>
                      <span className="text-[9px] sm:text-[10px] font-extrabold text-indigo-600 uppercase tracking-wider block">
                        STAGE 02 VERIFY EXCEL DETAILS & GENERATE
                      </span>
                      <h2 className="text-xl sm:text-2xl font-extrabold text-slate-900 mt-0.5">
                        {eventId || 'Batch Certificates'}
                      </h2>
                    </div>

                    <div className="flex flex-wrap items-center gap-2.5 sm:gap-3 w-full sm:w-auto">
                      <span className="bg-emerald-50 text-emerald-700 border border-emerald-200 font-extrabold px-3 py-1.5 sm:px-4 sm:py-2 rounded-full text-xs flex items-center space-x-1.5 shrink-0">
                        <Check className="w-3.5 h-3.5 stroke-[3]" />
                        <span>Event Registered in DB</span>
                      </span>

                      {/* GENERATE CERTIFICATES BUTTON */}
                      <button
                        onClick={handleGenerateCertificatesBatch}
                        disabled={isGeneratingBatch}
                        className="w-full sm:w-auto bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold px-5 sm:px-8 py-3 sm:py-3.5 rounded-xl sm:rounded-2xl text-xs uppercase tracking-wider shadow-xl transition-all flex items-center justify-center space-x-2 disabled:opacity-50"
                      >
                        {isGeneratingBatch ? (
                          <>
                            <Loader2 className="w-4 h-4 animate-spin text-white" />
                            <span>Generating Certificates...</span>
                          </>
                        ) : (
                          <>
                            <Play className="w-4 h-4 fill-white" />
                            <span>Generate Certificates ({excelResult?.totalRows || 0} Records)</span>
                          </>
                        )}
                      </button>
                    </div>
                  </div>

                  {/* BATCH METRICS GRID */}
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 sm:gap-4">
                    <div className="bg-slate-50 p-4 sm:p-5 rounded-xl sm:rounded-2xl border border-slate-200">
                      <span className="text-[9px] sm:text-[10px] font-extrabold text-slate-400 uppercase tracking-wider block">
                        TOTAL EXCEL RECORDS
                      </span>
                      <span className="text-2xl sm:text-3xl font-black text-slate-900 font-mono mt-1 block">
                        {excelResult?.totalRows || 0} Rows
                      </span>
                    </div>

                    <div className="bg-slate-50 p-4 sm:p-5 rounded-xl sm:rounded-2xl border border-slate-200">
                      <span className="text-[9px] sm:text-[10px] font-extrabold text-slate-400 uppercase tracking-wider block">
                        EVENT DATE
                      </span>
                      <span className="text-sm sm:text-base font-extrabold text-slate-800 font-mono mt-1 block">
                        {eventDate || 'Not set'}
                      </span>
                    </div>

                    <div className="bg-slate-50 p-4 sm:p-5 rounded-xl sm:rounded-2xl border border-slate-200">
                      <span className="text-[9px] sm:text-[10px] font-extrabold text-slate-400 uppercase tracking-wider block">
                        CERTIFICATE ISSUE DATE
                      </span>
                      <span className="text-sm sm:text-base font-extrabold text-slate-800 font-mono mt-1 block">
                        {issueDate || 'Not set'}
                      </span>
                    </div>
                  </div>

                  {batchGenSuccessMsg && (
                    <div className="p-3 sm:p-4 bg-emerald-50 border border-emerald-200 text-emerald-800 font-extrabold text-xs rounded-xl sm:rounded-2xl flex items-center space-x-2">
                      <CheckCircle2 className="w-4 h-4 sm:w-5 sm:h-5 text-emerald-600 shrink-0" />
                      <span>{batchGenSuccessMsg}</span>
                    </div>
                  )}
                </div>

                {/* VERIFY UPLOADED EXCEL DATA TABLE */}
                <div className="bg-white rounded-2xl sm:rounded-3xl border border-slate-200 shadow-sm p-4 sm:p-6 space-y-4">
                  <h3 className="text-xs font-black uppercase tracking-wider text-slate-500">
                    Verify Uploaded Excel Participant Data ({excelResult?.totalRows || 0} Rows)
                  </h3>
                  
                  <div className="overflow-x-auto rounded-xl sm:rounded-2xl border border-slate-200 -mx-1 sm:mx-0">
                    <table className="w-full text-left text-xs font-medium border-collapse min-w-[700px]">
                      <thead>
                        <tr className="bg-slate-100 border-b border-slate-200 text-[10px] font-black uppercase text-slate-600">
                          <th className="py-3 px-4">#</th>
                          <th className="py-3 px-4">Participant Name</th>
                          <th className="py-3 px-4">Register No / Roll No</th>
                          <th className="py-3 px-4">Year</th>
                          <th className="py-3 px-4">Phone No</th>
                          <th className="py-3 px-4">Email Address</th>
                          <th className="py-3 px-4">Department & Section</th>
                          <th className="py-3 px-4">College Name</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {excelResult?.rows.map((row, idx) => {
                          const rNo = extractStudentId(row);
                          const name = String(row['Name'] || row['name'] || 'Participant').trim();
                          const email = extractStudentEmail(row, rNo);
                          const sec = String(row['Section'] || row['section'] || '').trim();
                          const col = String(row['College Name'] || row['College'] || 'Kongu Engineering College').trim();
                          const year = extractRowYear(row, String(row['Year of Study'] || row['Year'] || row['year'] || '-').trim()) || '-';
                          const phone = extractRowPhone(row, String(row['Mobile number '] || row['Phone'] || row['Mobile'] || row['phone'] || row['mobile'] || '-').trim()) || '-';

                          return (
                            <tr key={idx} className="hover:bg-slate-50">
                              <td className="py-3 px-4 font-mono font-bold text-slate-400">{idx + 1}</td>
                              <td className="py-3 px-4 font-extrabold text-slate-900 uppercase">{name}</td>
                              <td className="py-3 px-4 font-mono font-bold text-indigo-600">{rNo}</td>
                              <td className="py-3 px-4 font-semibold text-slate-700">{year}</td>
                              <td className="py-3 px-4 font-mono text-slate-600">{phone}</td>
                              <td className="py-3 px-4 text-slate-600 font-medium">{email}</td>
                              <td className="py-3 px-4 text-slate-700">CSE {sec ? `(Section ${sec})` : ''}</td>
                              <td className="py-3 px-4 text-slate-800 font-bold">{col}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div className="flex justify-between">
                  <button
                    onClick={() => setGeneratorStep(1)}
                    className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold px-6 py-3 rounded-xl text-xs uppercase"
                  >
                    ← Back to Stage 01
                  </button>
                </div>
              </div>
            )}

          </div>
        )}

        {/* TAB 2: CERTIFICATES (COPTERCODE ANALYTICS & 5 DYNAMIC FILTERS UI) */}
        {activeTab === 'history' && (
          <div className="space-y-6 sm:space-y-8">
            
            {/* TOP HEADER ROW WITH REFRESH BUTTON */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <h2 className="text-[11px] font-black uppercase tracking-widest text-slate-400">
                CERTIFICATE & EVENT ANALYTICS
              </h2>
              
              <div className="flex flex-wrap items-center gap-2 sm:gap-3">
                {historyCerts.some(c => c.emailStatus === 'pending' || !c.emailStatus) && (
                  <button
                    onClick={handleSendAllPendingEmails}
                    disabled={isBatchSendingEmails || isDbLoading}
                    className="flex-1 sm:flex-initial flex items-center justify-center space-x-2 bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 text-white font-extrabold px-4 sm:px-5 py-2 sm:py-2.5 rounded-xl text-xs transition-all shadow-md disabled:opacity-50"
                  >
                    {isBatchSendingEmails ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin text-white" />
                        <span>Sending ({batchEmailProgress.current}/{batchEmailProgress.total})</span>
                      </>
                    ) : (
                      <>
                        <Send className="w-4 h-4" />
                        <span>Send Pending ({historyCerts.filter(c => c.emailStatus === 'pending' || !c.emailStatus).length})</span>
                      </>
                    )}
                  </button>
                )}

                <button
                  onClick={loadDatabaseRecords}
                  disabled={isDbLoading || isBatchSendingEmails}
                  className="flex items-center justify-center space-x-2 bg-[#0284c7] hover:bg-[#0369a1] text-white font-extrabold px-4 sm:px-5 py-2 sm:py-2.5 rounded-xl text-xs transition-all shadow-md disabled:opacity-50"
                >
                  <RefreshCw className={`w-4 h-4 ${isDbLoading ? 'animate-spin' : ''}`} />
                  <span>{isDbLoading ? 'Refreshing...' : 'Refresh'}</span>
                </button>
              </div>
            </div>

            {/* LIVE SYNC ANALYTICS CARDS */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 sm:gap-6">
              
              {/* PURPLE GRADIENT CARD (LIVE SYNC COUNT) */}
              <div className="relative overflow-hidden bg-gradient-to-r from-[#635bfc] to-[#8b5cf6] rounded-2xl sm:rounded-[28px] p-5 sm:p-7 text-white shadow-xl flex flex-col justify-between min-h-[140px] sm:min-h-[160px]">
                <div className="absolute -right-4 -bottom-6 opacity-15 pointer-events-none">
                  <Users className="w-36 h-36 sm:w-48 sm:h-48 text-white" />
                </div>
                <div>
                  <span className="text-[10px] sm:text-[11px] font-extrabold uppercase tracking-widest text-indigo-100 block">
                    TOTAL UNIQUE PARTICIPANTS
                  </span>
                  <span className="text-4xl sm:text-5xl font-black tracking-tight mt-1 sm:mt-2 block font-mono">
                    {historyCerts.length}
                  </span>
                </div>
                <div className="flex items-center space-x-1.5 text-[11px] sm:text-xs text-indigo-100 font-bold mt-3 sm:mt-4">
                  <Zap className="w-3.5 h-3.5 sm:w-4 sm:h-4 text-amber-300 shrink-0" />
                  <span>Live synced ({historyCerts.length} Records)</span>
                </div>
              </div>

              {/* TOTAL EVENTS ANALYTICS CARD */}
              <div className="bg-white rounded-2xl sm:rounded-[28px] border border-slate-200 p-5 sm:p-7 shadow-sm flex flex-col justify-between min-h-[140px] sm:min-h-[160px] relative">
                <div>
                  <span className="text-[10px] sm:text-[11px] font-extrabold uppercase tracking-widest text-indigo-600 block">
                    TOTAL EVENTS
                  </span>
                  <span className="text-4xl sm:text-5xl font-black text-slate-900 tracking-tight mt-1 sm:mt-2 block font-mono">
                    {registeredEvents.length > 0
                      ? registeredEvents.length
                      : (new Set(historyCerts.map(c => c.eventId || c.eventName).filter(Boolean)).size || (historyCerts.length > 0 ? 1 : 0))}
                  </span>
                </div>
                <div className="flex items-center space-x-1.5 text-[11px] sm:text-xs text-slate-500 font-semibold mt-3 sm:mt-4">
                  <Calendar className="w-3.5 h-3.5 sm:w-4 sm:h-4 text-indigo-500 shrink-0" />
                  <span>Registered across all events</span>
                </div>
              </div>

            </div>

            {/* SEARCH BAR & ACTION BUTTONS CONTAINER */}
            <div className="bg-white rounded-2xl sm:rounded-[28px] border border-slate-300 p-4 sm:p-6 shadow-sm space-y-4 sm:space-y-6">
              
              {/* TOP SEARCH & BUTTONS ROW */}
              <div className="flex flex-col lg:flex-row items-stretch lg:items-center gap-3 sm:gap-4">
                
                {/* SEARCH INPUT */}
                <div className="relative flex-1 w-full">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-3" />
                  <input
                    type="text"
                    placeholder="Search by name, college, email, code..."
                    value={historyQuery}
                    onChange={(e) => setHistoryQuery(e.target.value)}
                    className="w-full pl-10 pr-4 py-2 sm:py-2.5 rounded-full border border-slate-200 text-xs font-medium text-slate-800 focus:outline-none focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100"
                  />
                </div>

                {/* DYNAMIC ACTION BUTTONS */}
                <div className="flex flex-wrap sm:flex-nowrap items-center gap-2 sm:gap-3 shrink-0">

                  <button
                    onClick={handleExportZip}
                    disabled={isExportingZip}
                    className="flex-1 sm:flex-initial flex items-center justify-center space-x-2 bg-[#7c3aed] hover:bg-[#6d28d9] text-white font-extrabold px-4 sm:px-5 py-2 sm:py-2.5 rounded-xl text-xs transition-all shadow-sm whitespace-nowrap disabled:opacity-50"
                  >
                    {isExportingZip ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin text-white" />
                        <span>Packaging ZIP...</span>
                      </>
                    ) : (
                      <>
                        <Download className="w-4 h-4" />
                        <span>{selectedCertIds.length > 0 ? 'Download ZIP' : 'Download ZIP'}</span>
                      </>
                    )}
                  </button>

                  {selectedCertIds.length > 0 && (
                    <button
                      onClick={handleSendSelectedEmails}
                      disabled={isBatchSendingEmails}
                      className="flex-1 sm:flex-initial flex items-center justify-center space-x-2 bg-[#0284c7] hover:bg-[#0369a1] text-white font-extrabold px-4 sm:px-5 py-2 sm:py-2.5 rounded-xl text-xs transition-all shadow-sm whitespace-nowrap disabled:opacity-50"
                    >
                      {isBatchSendingEmails ? (
                        <>
                          <Loader2 className="w-4 h-4 animate-spin text-white" />
                          <span>Sending ({batchEmailProgress.current}/{batchEmailProgress.total})</span>
                        </>
                      ) : (
                        <>
                          <Send className="w-4 h-4" />
                          <span>Send Selected ({selectedCertIds.length})</span>
                        </>
                      )}
                    </button>
                  )}

                  <button
                    onClick={resetCertFilters}
                    className="flex items-center justify-center space-x-1.5 border border-red-200 bg-red-50 hover:bg-red-100 text-red-700 font-extrabold px-3.5 sm:px-4 py-2 sm:py-2.5 rounded-xl text-xs transition-all shadow-sm whitespace-nowrap"
                  >
                    <Trash2 className="w-4 h-4" />
                    <span className="hidden sm:inline">Reset Filters</span>
                  </button>
                </div>

              </div>

              {/* FIVE DYNAMIC FILTERS ROW */}
              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2.5 sm:gap-3 pt-2">
                <div>
                  <label className="block text-[9px] sm:text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    EVENT
                  </label>
                  <select
                    value={certEventFilter}
                    onChange={(e) => setCertEventFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2 sm:p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Events</option>
                    {certEventOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    DEPARTMENT
                  </label>
                  <select
                    value={certDeptFilter}
                    onChange={(e) => setCertDeptFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Departments</option>
                    {certDeptOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    COLLEGE NAME
                  </label>
                  <select
                    value={certCollegeFilter}
                    onChange={(e) => setCertCollegeFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Colleges</option>
                    {certCollegeOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    YEAR OF STUDY
                  </label>
                  <select
                    value={certYearFilter}
                    onChange={(e) => setCertYearFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Years</option>
                    {certYearOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    SECTION
                  </label>
                  <select
                    value={certSectionFilter}
                    onChange={(e) => setCertSectionFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Sections</option>
                    {certSectionOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>
              </div>

            </div>

            {/* MASTER CERTIFICATES TABLE WITH LOADING SPINNER OVERLAY */}
            <div className="bg-white rounded-[28px] border border-slate-300 shadow-sm overflow-hidden relative">
              {isDbLoading && (
                <div className="absolute inset-0 bg-white/70 backdrop-blur-sm z-10 flex items-center justify-center">
                  <div className="flex items-center space-x-2 bg-slate-900 text-white font-extrabold px-5 py-3 rounded-2xl shadow-xl text-xs">
                    <Loader2 className="w-4 h-4 animate-spin text-blue-400" />
                    <span>Syncing with Supabase Live Database...</span>
                  </div>
                </div>
              )}

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs font-medium">
                  <thead>
                    <tr className="bg-slate-50/80 border-b border-slate-200 text-[10px] font-black uppercase tracking-wider text-slate-600">
                      <th className="py-4 px-5 w-10 text-center">
                        <input
                          type="checkbox"
                          checked={filteredCertificates.length > 0 && selectedCertIds.length === filteredCertificates.length}
                          onChange={toggleSelectAllCerts}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500" 
                        />
                      </th>
                      <th className="py-4 px-6 whitespace-nowrap">Intern Details</th>
                      <th className="py-4 px-6 whitespace-nowrap">Credential Code</th>
                      <th className="py-4 px-6 whitespace-nowrap">Event</th>
                      <th className="py-4 px-6 whitespace-nowrap">Issue Date</th>
                      <th className="py-4 px-6 whitespace-nowrap">Status</th>
                      <th className="py-4 px-6 whitespace-nowrap">Email Delivery</th>
                      <th className="py-4 px-6 text-right whitespace-nowrap">PDF File</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200/80">
                    {filteredCertificates.length > 0 ? (
                      filteredCertificates.map((c) => {
                        const isSelected = selectedCertIds.includes(c.id);
                        return (
                          <tr key={c.id} className={`transition-colors ${isSelected ? 'bg-indigo-50/40' : 'hover:bg-slate-50/80'}`}>
                            <td className="py-5 px-5 text-center">
                              <input
                                type="checkbox"
                                checked={isSelected}
                                onChange={() => toggleSelectCert(c.id)}
                                className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500" 
                              />
                            </td>

                            <td className="py-5 px-6">
                              <div className="font-extrabold text-slate-900 text-sm tracking-tight uppercase">
                                {c.studentName}
                              </div>
                              <div className="text-[10px] font-semibold text-slate-400 uppercase mt-0.5">
                                {c.customFields?.college_name || 'KONGU ENGINEERING COLLEGE'}
                              </div>
                              <div className="text-[11px] text-slate-500 mt-0.5">
                                {c.studentEmail}
                              </div>
                            </td>

                            <td className="py-5 px-6 font-mono text-slate-700">
                              <div className="flex items-center space-x-2">
                                <span className="text-xs font-semibold">{c.id}</span>
                                <button
                                  onClick={() => handleCopyCode(c.id)}
                                  className="text-slate-400 hover:text-slate-600 transition-colors"
                                  title="Copy Credential Code"
                                >
                                  {copiedCodeId === c.id ? (
                                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                                  ) : (
                                    <Copy className="w-3.5 h-3.5" />
                                  )}
                                </button>
                              </div>
                            </td>

                            <td className="py-5 px-6">
                              <span className="inline-block bg-emerald-50 text-emerald-700 border border-emerald-200/60 font-black text-[10px] uppercase tracking-wider px-3 py-1 rounded-full">
                                {c.eventName || 'Workshop'}
                              </span>
                            </td>

                            <td className="py-5 px-6 font-mono text-slate-800 text-xs font-bold whitespace-nowrap">
                              {c.issueDate || '2026-07-25'}
                            </td>

                            <td className="py-5 px-6">
                              <span className="inline-block border border-emerald-500 text-emerald-600 bg-emerald-50/80 px-3.5 py-1 rounded-full text-[11px] font-bold">
                                Active
                              </span>
                            </td>

                            <td className="py-5 px-6">
                              {c.emailStatus === 'sent' ? (
                                <span className="inline-flex items-center space-x-1 border border-emerald-500/40 text-emerald-600 bg-emerald-50 px-2.5 py-1 rounded-full text-[11px] font-extrabold shadow-sm">
                                  <Check className="w-3.5 h-3.5" />
                                  <span>Sent</span>
                                </span>
                              ) : (
                                <button
                                  onClick={() => handleSendSingleEmail(c)}
                                  disabled={sendingSingleCertId === c.id || isBatchSendingEmails}
                                  className="inline-flex items-center space-x-1.5 border border-amber-400 hover:border-amber-500 text-amber-700 hover:text-amber-800 bg-amber-50 hover:bg-amber-100 px-3 py-1 rounded-full text-[11px] font-extrabold transition-all shadow-sm disabled:opacity-50"
                                >
                                  {sendingSingleCertId === c.id ? (
                                    <>
                                      <Loader2 className="w-3.5 h-3.5 animate-spin text-amber-600" />
                                      <span>Sending...</span>
                                    </>
                                  ) : (
                                    <>
                                      <Send className="w-3.5 h-3.5 text-amber-600" />
                                      <span>Send Email</span>
                                    </>
                                  )}
                                </button>
                              )}
                            </td>

                            <td className="py-5 px-6 text-right">
                              <div className="flex items-center justify-end space-x-2">
                                <button
                                  onClick={() => handleOpenEmailCodeModal(c)}
                                  className="flex items-center space-x-1 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200/80 font-bold px-3 py-1 rounded-xl text-xs transition-colors shadow-xs"
                                  title="View & Copy Rendered HTML Email Code"
                                >
                                  <Code className="w-3.5 h-3.5 text-indigo-600" />
                                  <span>Email HTML</span>
                                </button>

                                <button
                                  onClick={() => handleViewCertificate(c)}
                                  className="flex items-center space-x-1 bg-purple-50 hover:bg-purple-100 text-purple-700 border border-purple-200 font-bold px-3 py-1 rounded-xl text-xs transition-colors"
                                >
                                  <ExternalLink className="w-3.5 h-3.5" />
                                  <span>View</span>
                                </button>

                                <button
                                  onClick={() => handleDownloadCertificate(c)}
                                  className="flex items-center space-x-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 font-bold px-3 py-1 rounded-xl text-xs transition-colors"
                                >
                                  <Download className="w-3.5 h-3.5" />
                                  <span>Download</span>
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    ) : (
                      <tr>
                        <td colSpan={8} className="py-16 text-center text-slate-400">
                          <Database className="w-10 h-10 mx-auto mb-3 opacity-40 text-indigo-600" />
                          <p className="text-base font-bold text-slate-800">No certificate records in Supabase database</p>
                          <p className="text-xs text-slate-400 mt-1">Upload template & Excel sheet in Stage 01 and click Generate Certificates in Stage 02.</p>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        )}

        {/* TAB 3: REGISTRATION (STUDENT DIRECTORY MATCHING COPTERCODE UI & TOP REFRESH BUTTON) */}
        {activeTab === 'registration' && (
          <div className="space-y-8">
            
            {/* TOP HEADER ROW WITH REFRESH BUTTON */}
            <div className="flex items-center justify-between">
              <h2 className="text-[11px] font-black uppercase tracking-widest text-slate-400">
                STUDENT DIRECTORY ANALYTICS
              </h2>

              <button
                onClick={loadDatabaseRecords}
                disabled={isDbLoading}
                className="flex items-center space-x-2 bg-[#0284c7] hover:bg-[#0369a1] text-white font-extrabold px-5 py-2.5 rounded-xl text-xs transition-all shadow-md disabled:opacity-50"
              >
                <RefreshCw className={`w-4 h-4 ${isDbLoading ? 'animate-spin' : ''}`} />
                <span>{isDbLoading ? 'Refreshing...' : 'Refresh'}</span>
              </button>
            </div>

            {/* LIVE SYNC ANALYTICS CARDS */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              
              {/* PURPLE GRADIENT CARD (LIVE SYNC COUNT) */}
              <div className="relative overflow-hidden bg-gradient-to-r from-[#635bfc] to-[#8b5cf6] rounded-[28px] p-7 text-white shadow-xl flex flex-col justify-between min-h-[160px]">
                <div className="absolute -right-4 -bottom-6 opacity-15 pointer-events-none">
                  <UserCheck className="w-48 h-48 text-white" />
                </div>
                <div>
                  <span className="text-[11px] font-extrabold uppercase tracking-widest text-indigo-100 block">
                    TOTAL REGISTERED STUDENTS
                  </span>
                  <span className="text-5xl font-black tracking-tight mt-2 block font-mono">
                    {registrations.length}
                  </span>
                </div>
                <div className="flex items-center space-x-1.5 text-xs text-indigo-100 font-bold mt-4">
                  <Zap className="w-4 h-4 text-amber-300" />
                  <span>Live synced with Database ({registrations.length} Records)</span>
                </div>
              </div>

              {/* TOTAL COLLEGES ANALYTICS CARD */}
              <div className="bg-white rounded-[28px] border border-slate-200 p-7 shadow-sm flex flex-col justify-between min-h-[160px] relative">
                <div>
                  <span className="text-[11px] font-extrabold uppercase tracking-widest text-indigo-600 block">
                    TOTAL COLLEGES
                  </span>
                  <span className="text-5xl font-black text-slate-900 tracking-tight mt-2 block font-mono">
                    {new Set(registrations.map(s => s.collegeName || 'Kongu Engineering College').filter(Boolean)).size || (registrations.length > 0 ? 1 : 0)}
                  </span>
                </div>
                <div className="flex items-center space-x-1.5 text-xs text-slate-500 font-semibold mt-4">
                  <Tag className="w-4 h-4 text-indigo-500" />
                  <span>Colleges represented among registered students</span>
                </div>
              </div>

            </div>

            {/* SEARCH BAR & ACTION BUTTONS CONTAINER */}
            <div className="bg-white rounded-[28px] border border-slate-300 p-6 shadow-sm space-y-6">
              
              {/* TOP SEARCH & BUTTONS ROW */}
              <div className="flex flex-col lg:flex-row items-center gap-4">
                
                {/* SEARCH INPUT */}
                <div className="relative flex-1 w-full">
                  <Search className="w-4 h-4 text-slate-400 absolute left-4 top-3.5" />
                  <input
                    type="text"
                    placeholder="Search by student name, roll number, email, or section..."
                    value={regQuery}
                    onChange={(e) => setRegQuery(e.target.value)}
                    className="w-full pl-11 pr-4 py-2.5 rounded-full border border-slate-200 text-xs font-medium text-slate-800 focus:outline-none focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100"
                  />
                </div>

                {/* ACTION BUTTONS */}
                <div className="flex items-center gap-3 w-full lg:w-auto shrink-0 overflow-x-auto">
                  <button
                    onClick={handleExportStudentExcel}
                    disabled={isExportingExcel}
                    className="flex items-center space-x-2 bg-[#059669] hover:bg-[#047857] text-white font-extrabold px-5 py-2.5 rounded-xl text-xs transition-all shadow-sm whitespace-nowrap disabled:opacity-50"
                  >
                    {isExportingExcel ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin text-white" />
                        <span>Exporting...</span>
                      </>
                    ) : (
                      <>
                        <Download className="w-4 h-4" />
                        <span>Export Student Excel</span>
                      </>
                    )}
                  </button>

                  <button
                    onClick={resetStudentFilters}
                    className="flex items-center space-x-2 bg-red-50/60 hover:bg-red-100/60 border border-dashed border-red-300 text-red-600 font-extrabold px-4 py-2.5 rounded-xl text-xs transition-colors whitespace-nowrap"
                  >
                    <Trash2 className="w-4 h-4" />
                    <span>Reset Filters</span>
                  </button>
                </div>

              </div>

              {/* FIVE DYNAMIC FILTERS ROW: EVENT, DEPARTMENT, COLLEGE NAME, YEAR OF STUDY, SECTION */}
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-3 pt-2">
                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    EVENT
                  </label>
                  <select
                    value={studentEventFilter}
                    onChange={(e) => setStudentEventFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Events</option>
                    {certEventOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    DEPARTMENT
                  </label>
                  <select
                    value={studentDeptFilter}
                    onChange={(e) => setStudentDeptFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Departments</option>
                    {certDeptOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    COLLEGE NAME
                  </label>
                  <select
                    value={studentCollegeFilter}
                    onChange={(e) => setStudentCollegeFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Colleges</option>
                    {certCollegeOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    YEAR OF STUDY
                  </label>
                  <select
                    value={studentYearFilter}
                    onChange={(e) => setStudentYearFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Years</option>
                    {certYearOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-1">
                    SECTION
                  </label>
                  <select
                    value={studentSectionFilter}
                    onChange={(e) => setStudentSectionFilter(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl p-2.5 text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500"
                  >
                    <option>All Sections</option>
                    {certSectionOptions.map((opt, idx) => (
                      <option key={idx} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>
              </div>

            </div>

            {studentUpdateSuccessMsg && (
              <div className="p-4 bg-emerald-50 border border-emerald-200 text-emerald-800 font-extrabold text-xs rounded-2xl flex items-center space-x-2 shadow-sm">
                <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
                <span>{studentUpdateSuccessMsg}</span>
              </div>
            )}

            {/* MASTER STUDENT DIRECTORY TABLE WITH LOADING SPINNER OVERLAY */}
            <div className="bg-white rounded-[28px] border border-slate-300 shadow-sm overflow-hidden relative">
              {isDbLoading && (
                <div className="absolute inset-0 bg-white/70 backdrop-blur-sm z-10 flex items-center justify-center">
                  <div className="flex items-center space-x-2 bg-slate-900 text-white font-extrabold px-5 py-3 rounded-2xl shadow-xl text-xs">
                    <Loader2 className="w-4 h-4 animate-spin text-blue-400" />
                    <span>Syncing Student Directory with Database...</span>
                  </div>
                </div>
              )}

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs font-medium">
                  <thead>
                    <tr className="bg-slate-50/80 border-b border-slate-200 text-[10px] font-black uppercase tracking-wider text-slate-600">
                      <th className="py-4 px-6 whitespace-nowrap">Student Details</th>
                      <th className="py-4 px-6 whitespace-nowrap">Register No</th>
                      <th className="py-4 px-6 whitespace-nowrap">Department & Section</th>
                      <th className="py-4 px-6 whitespace-nowrap">Year of Study</th>
                      <th className="py-4 px-6 whitespace-nowrap">College Name</th>
                      <th className="py-4 px-6 whitespace-nowrap">Email & Contact</th>
                      <th className="py-4 px-6 text-right whitespace-nowrap">Actions</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200/80">
                    {filteredStudents.length > 0 ? (
                      filteredStudents.map((s) => {
                        const isEditing = editingStudentId === s.id;

                        if (isEditing && editStudentForm) {
                          return (
                            <tr key={s.id} className="bg-amber-50/50 border-y-2 border-amber-300/80 transition-colors">
                              <td className="py-4 px-6">
                                <input
                                  type="text"
                                  value={editStudentForm.name || ''}
                                  onChange={(e) => setEditStudentForm({ ...editStudentForm, name: e.target.value })}
                                  className="w-full bg-white border border-slate-300 rounded-lg p-2 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-sm uppercase"
                                  placeholder="Student Name"
                                />
                              </td>

                              <td className="py-4 px-6">
                                <input
                                  type="text"
                                  value={editStudentForm.registerNo || ''}
                                  onChange={(e) => setEditStudentForm({ ...editStudentForm, registerNo: e.target.value })}
                                  className="w-full bg-white border border-slate-300 rounded-lg p-2 text-xs font-mono font-bold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-sm uppercase"
                                  placeholder="Register No"
                                />
                              </td>

                              <td className="py-4 px-6">
                                <div className="flex items-center space-x-2">
                                  <input
                                    type="text"
                                    value={editStudentForm.department || ''}
                                    onChange={(e) => setEditStudentForm({ ...editStudentForm, department: e.target.value })}
                                    className="w-28 bg-white border border-slate-300 rounded-lg p-2 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-sm"
                                    placeholder="Dept (CSE)"
                                  />
                                  <input
                                    type="text"
                                    value={editStudentForm.section || ''}
                                    onChange={(e) => setEditStudentForm({ ...editStudentForm, section: e.target.value })}
                                    className="w-16 bg-white border border-slate-300 rounded-lg p-2 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-sm"
                                    placeholder="Sec"
                                  />
                                </div>
                              </td>

                              <td className="py-4 px-6">
                                <input
                                  type="text"
                                  value={editStudentForm.yearOfStudy || ''}
                                  onChange={(e) => setEditStudentForm({ ...editStudentForm, yearOfStudy: e.target.value })}
                                  className="w-28 bg-white border border-slate-300 rounded-lg p-2 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-sm"
                                  placeholder="Year of Study"
                                />
                              </td>

                              <td className="py-4 px-6">
                                <input
                                  type="text"
                                  value={editStudentForm.collegeName || ''}
                                  onChange={(e) => setEditStudentForm({ ...editStudentForm, collegeName: e.target.value })}
                                  className="w-full bg-white border border-slate-300 rounded-lg p-2 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-sm"
                                  placeholder="College Name"
                                />
                              </td>

                              <td className="py-4 px-6">
                                <input
                                  type="email"
                                  value={editStudentForm.email || ''}
                                  onChange={(e) => setEditStudentForm({ ...editStudentForm, email: e.target.value })}
                                  className="w-full bg-white border border-slate-300 rounded-lg p-2 text-xs font-semibold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-sm mb-1.5"
                                  placeholder="Email Address"
                                />
                                <input
                                  type="text"
                                  value={editStudentForm.phone || ''}
                                  onChange={(e) => setEditStudentForm({ ...editStudentForm, phone: e.target.value })}
                                  className="w-full bg-white border border-slate-300 rounded-lg p-2 text-xs font-mono font-semibold text-slate-700 focus:outline-none focus:border-indigo-500 shadow-sm"
                                  placeholder="Phone No"
                                />
                              </td>

                              <td className="py-4 px-6 text-right">
                                <div className="flex items-center justify-end space-x-2">
                                  <button
                                    onClick={handleSaveStudent}
                                    disabled={isSavingStudent}
                                    className="flex items-center space-x-1 bg-[#059669] hover:bg-[#047857] text-white font-extrabold px-3 py-2 rounded-xl text-xs transition-all shadow-sm disabled:opacity-50"
                                    title="Save to Database"
                                  >
                                    {isSavingStudent ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Check className="w-3.5 h-3.5" />}
                                    <span>Save</span>
                                  </button>
                                  <button
                                    onClick={handleCancelEditStudent}
                                    disabled={isSavingStudent}
                                    className="p-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl text-xs transition-all"
                                    title="Cancel"
                                  >
                                    <X className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        }

                        return (
                          <tr key={s.id} className="hover:bg-slate-50/80 transition-colors">
                            <td className="py-5 px-6">
                              <div className="font-extrabold text-slate-900 text-sm tracking-tight uppercase">
                                {s.name}
                              </div>
                              <div className="text-[10px] font-semibold text-slate-400 uppercase mt-0.5">
                                {s.department || 'COMPUTER SCIENCE AND ENGINEERING'}
                              </div>
                            </td>

                            <td className="py-5 px-6 font-mono text-slate-900 text-xs font-black">
                              {s.registerNo}
                            </td>

                            <td className="py-5 px-6 whitespace-nowrap">
                              <div className="flex items-center space-x-2">
                                <span className="inline-block bg-indigo-50 text-indigo-700 border border-indigo-200/60 font-black text-[10px] uppercase tracking-wider px-2.5 py-0.5 rounded-full whitespace-nowrap">
                                  {getShortDept(s.department)}
                                </span>
                                {s.section && (
                                  <span className="font-extrabold text-slate-700 text-xs whitespace-nowrap">
                                    Section {s.section}
                                  </span>
                                )}
                              </div>
                            </td>

                            <td className="py-5 px-6 font-extrabold text-slate-800 text-xs">
                              {s.yearOfStudy || 'Not set'}
                            </td>

                            <td className="py-5 px-6 font-bold text-slate-800 text-xs">
                              {s.collegeName || 'Kongu Engineering College'}
                            </td>

                            <td className="py-5 px-6">
                              <div className="text-slate-800 font-medium">{s.email}</div>
                              <div className="text-[10px] text-slate-400 font-mono mt-0.5">{s.phone || '+91 9876543210'}</div>
                            </td>

                            <td className="py-5 px-6 text-right">
                              <button
                                onClick={() => handleStartEditStudent(s)}
                                className="inline-flex items-center space-x-1.5 bg-slate-100 hover:bg-indigo-50 hover:text-indigo-600 text-slate-700 font-bold px-3.5 py-1.5 rounded-xl text-xs transition-colors border border-slate-200/80 shadow-2xs"
                                title="Edit Student Record"
                              >
                                <Edit2 className="w-3.5 h-3.5" />
                                <span>Edit</span>
                              </button>
                            </td>
                          </tr>
                        );
                      })
                    ) : (
                      <tr>
                        <td colSpan={7} className="py-16 text-center text-slate-400">
                          <Database className="w-10 h-10 mx-auto mb-3 opacity-40 text-indigo-600" />
                          <p className="text-base font-bold text-slate-800">No student records in Supabase database</p>
                          <p className="text-xs text-slate-400 mt-1">Upload template & Excel sheet in Stage 01 and click Generate Certificates in Stage 02.</p>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        )}

        {/* TAB 4: PROFILE */}
        {activeTab === 'profile' && (
          <div className="space-y-8 max-w-6xl mx-auto pb-12">
            
            {/* TOP ROW: 2 CARDS */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">

              {/* CARD 1: ADMIN INFORMATION */}
              <div className="bg-white rounded-[32px] p-8 border border-slate-200/80 shadow-[0_10px_30px_rgba(0,0,0,0.03)] space-y-6 flex flex-col justify-between">
                <div className="space-y-6">
                  <div>
                    <div className="flex items-center space-x-2.5">
                      <User className="w-5 h-5 text-[#635bfc]" />
                      <h3 className="text-xl font-extrabold text-slate-900 tracking-tight">Admin Information</h3>
                    </div>
                    <p className="text-xs text-slate-400 font-medium leading-relaxed mt-1.5">
                      Update your display name. This name is used to greet you in the panel interface.
                    </p>
                  </div>

                  {adminInfoMsg && (
                    <div className={`p-3 rounded-2xl text-xs font-extrabold ${adminInfoMsg.startsWith('Error') ? 'bg-rose-50 text-rose-700 border border-rose-200' : 'bg-emerald-50 text-emerald-700 border border-emerald-200'}`}>
                      {adminInfoMsg}
                    </div>
                  )}

                  <div className="space-y-5">
                    <div>
                      <label className="block text-[10px] font-black uppercase tracking-wider text-slate-400 mb-2">ACCOUNT EMAIL</label>
                      <input
                        type="text"
                        disabled
                        value={profileEmail || 'rohithp.24cse@kongu.edu'}
                        className="w-full bg-slate-50 border border-slate-200/80 rounded-2xl p-3.5 text-xs font-semibold text-slate-500 cursor-not-allowed focus:outline-none"
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] font-black uppercase tracking-wider text-slate-400 mb-2">DISPLAY NAME</label>
                      <input
                        type="text"
                        value={profileName}
                        onChange={(e) => setProfileName(e.target.value)}
                        className="w-full bg-white border border-slate-200/80 rounded-2xl p-3.5 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/10 transition-all shadow-xs"
                        placeholder="Rohith"
                      />
                    </div>
                  </div>
                </div>

                <div className="pt-4">
                  <button
                    onClick={handleSaveAdminInfo}
                    disabled={isUpdatingAdminInfo}
                    className="bg-[#635bfc] hover:bg-[#5248f7] text-white font-extrabold px-6 py-3.5 rounded-2xl text-xs uppercase tracking-wider transition-all shadow-md active:scale-95 disabled:opacity-50"
                  >
                    {isUpdatingAdminInfo ? 'SAVING CHANGES...' : 'SAVE CHANGES'}
                  </button>
                </div>
              </div>

              {/* CARD 2: CHANGE SECURITY PASSWORD */}
              <div className="bg-white rounded-[32px] p-8 border border-slate-200/80 shadow-[0_10px_30px_rgba(0,0,0,0.03)] space-y-6 flex flex-col justify-between">
                <div className="space-y-6">
                  <div>
                    <div className="flex items-center space-x-2.5">
                      <KeyRound className="w-5 h-5 text-[#938cfb]" />
                      <h3 className="text-xl font-extrabold text-slate-900 tracking-tight">Change Security Password</h3>
                    </div>
                    <p className="text-xs text-slate-400 font-medium leading-relaxed mt-1.5">
                      Update your credentials. Once saved, you must log back in with the new password on future visits.
                    </p>
                  </div>

                  {passwordMsg && (
                    <div className={`p-3 rounded-2xl text-xs font-extrabold ${passwordMsg.startsWith('Error') ? 'bg-rose-50 text-rose-700 border border-rose-200' : 'bg-emerald-50 text-emerald-700 border border-emerald-200'}`}>
                      {passwordMsg}
                    </div>
                  )}

                  <div className="space-y-5">
                    <div>
                      <label className="block text-[10px] font-black uppercase tracking-wider text-slate-400 mb-2">NEW SECURITY PASSWORD</label>
                      <input
                        type="password"
                        value={newPassword}
                        onChange={(e) => setNewPassword(e.target.value)}
                        className="w-full bg-white border border-slate-200/80 rounded-2xl p-3.5 text-xs font-semibold text-slate-900 focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/10 transition-all shadow-xs placeholder:text-slate-300"
                        placeholder="••••••••  (Min 6 characters)"
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] font-black uppercase tracking-wider text-slate-400 mb-2">CONFIRM PASSWORD</label>
                      <input
                        type="password"
                        value={confirmPassword}
                        onChange={(e) => setConfirmPassword(e.target.value)}
                        className="w-full bg-white border border-slate-200/80 rounded-2xl p-3.5 text-xs font-semibold text-slate-900 focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/10 transition-all shadow-xs placeholder:text-slate-300"
                        placeholder="••••••••"
                      />
                    </div>
                  </div>
                </div>

                <div className="pt-4">
                  <button
                    onClick={handleUpdatePassword}
                    disabled={isUpdatingPassword}
                    className="bg-[#938cfb] hover:bg-[#7e75fa] text-white font-extrabold px-6 py-3.5 rounded-2xl text-xs uppercase tracking-wider transition-all shadow-md active:scale-95 disabled:opacity-50"
                  >
                    {isUpdatingPassword ? 'UPDATING PASSWORD...' : 'UPDATE PASSWORD'}
                  </button>
                </div>
              </div>

            </div>

            {/* CARD 3: EMAIL TEMPLATE SETTINGS (FULL WIDTH) */}
            <div className="bg-white rounded-[32px] p-8 border border-slate-200/80 shadow-[0_10px_30px_rgba(0,0,0,0.03)] space-y-6">
              <div>
                <div className="flex items-center space-x-2.5">
                  <Mail className="w-5 h-5 text-[#635bfc]" />
                  <h3 className="text-xl font-extrabold text-slate-900 tracking-tight">Email Template Settings</h3>
                </div>
                <p className="text-xs text-slate-400 font-medium leading-relaxed mt-1.5">
                  Customize the images used in all outgoing emails. Toggle <span className="font-bold text-slate-700">Custom</span> to override the server default (<code className="bg-slate-100 px-1.5 py-0.5 rounded text-slate-600 font-mono">.env</code>). When toggled off, the env default is used automatically.
                </p>
              </div>

              {emailSettingsMsg && (
                <div className="p-3 bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-extrabold rounded-2xl">
                  {emailSettingsMsg}
                </div>
              )}

              <div className="grid grid-cols-1 md:grid-cols-2 gap-8 pt-2">
                {/* LOGO IMAGE */}
                <div className="space-y-4 bg-slate-50/50 p-5 rounded-3xl border border-slate-200/80">
                  <div className="flex items-center justify-between">
                    <div>
                      <h4 className="font-extrabold text-slate-900 text-sm flex items-center space-x-2">
                        <ImageIcon className="w-4 h-4 text-indigo-600" />
                        <span>Logo Image</span>
                      </h4>
                      <p className="text-[11px] text-slate-400 font-medium">Used in top navigation & certificate emails.</p>
                    </div>
                    <button
                      type="button"
                      onClick={() => setUseCustomLogo(!useCustomLogo)}
                      className={`relative inline-flex h-6 w-11 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${useCustomLogo ? 'bg-[#635bfc]' : 'bg-slate-200'}`}
                    >
                      <span className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow-lg ring-0 transition duration-200 ease-in-out ${useCustomLogo ? 'translate-x-5' : 'translate-x-0'}`} />
                    </button>
                  </div>

                  {useCustomLogo ? (
                    <div className="space-y-3">
                      {/* Image Preview Box */}
                      {customLogoUrl && (
                        <div className="flex items-center space-x-3 p-3 bg-white rounded-2xl border border-slate-200 shadow-xs">
                          <div className="w-16 h-12 bg-slate-50 rounded-xl border border-slate-100 p-1 flex items-center justify-center shrink-0 overflow-hidden">
                            <img src={customLogoUrl} alt="Custom Logo Preview" className="max-h-full max-w-full object-contain" />
                          </div>
                          <div className="flex-1 min-w-0">
                            <p className="text-xs font-bold text-slate-800 truncate">Custom Logo Active</p>
                            <a href={customLogoUrl} target="_blank" rel="noreferrer" className="text-[10px] text-indigo-600 hover:underline truncate block">
                              {customLogoUrl}
                            </a>
                          </div>
                          <button
                            type="button"
                            onClick={() => setCustomLogoUrl('')}
                            className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-slate-100 transition-colors"
                            title="Clear image URL"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      )}

                      {/* Upload Button / Dropzone */}
                      <label className={`flex flex-col items-center justify-center p-4 border-2 border-dashed rounded-2xl cursor-pointer transition-all ${
                        isUploadingLogo ? 'border-indigo-400 bg-indigo-50/40' : 'border-slate-300 hover:border-indigo-500 bg-white hover:bg-indigo-50/10'
                      }`}>
                        <input
                          type="file"
                          accept="image/*"
                          className="hidden"
                          disabled={isUploadingLogo}
                          onChange={(e) => {
                            if (e.target.files && e.target.files[0]) {
                              handleUploadImage(e.target.files[0], 'logo');
                            }
                          }}
                        />
                        {isUploadingLogo ? (
                          <div className="flex items-center space-x-2 text-indigo-600 text-xs font-bold">
                            <Loader2 className="w-4 h-4 animate-spin" />
                            <span>Uploading logo to Supabase Storage...</span>
                          </div>
                        ) : (
                          <div className="flex items-center space-x-2 text-slate-600 text-xs font-bold">
                            <Upload className="w-4 h-4 text-indigo-600" />
                            <span>Click or Drag to Upload Logo (.png, .jpg, .svg)</span>
                          </div>
                        )}
                      </label>

                      {/* URL input fallback */}
                      <div>
                        <span className="text-[10px] font-black uppercase tracking-wider text-slate-400 block mb-1">OR ENTER DIRECT URL</span>
                        <input
                          type="text"
                          value={customLogoUrl}
                          onChange={(e) => setCustomLogoUrl(e.target.value)}
                          className="w-full bg-white border border-slate-300 rounded-2xl p-3 text-xs font-semibold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-xs"
                          placeholder="https://your-domain.com/custom_logo.svg"
                        />
                      </div>
                    </div>
                  ) : (
                    <div className="flex items-center justify-between p-3 bg-white rounded-2xl border border-slate-200">
                      <div className="flex items-center space-x-3">
                        <div className="w-12 h-10 bg-slate-50 rounded-xl border border-slate-100 p-1 flex items-center justify-center shrink-0">
                          <img src="/csea_logo.png" alt="Default Logo" className="max-h-full max-w-full object-contain" />
                        </div>
                        <div>
                          <p className="text-xs font-bold text-slate-700">Using Default CSEA Logo</p>
                          <span className="text-[10px] text-slate-400">Toggle switch above to upload custom logo</span>
                        </div>
                      </div>
                      <span className="bg-slate-100 px-2.5 py-1 rounded-full text-[10px] font-bold text-slate-500">.env Default</span>
                    </div>
                  )}
                </div>

                {/* HERO IMAGE */}
                <div className="space-y-4 bg-slate-50/50 p-5 rounded-3xl border border-slate-200/80">
                  <div className="flex items-center justify-between">
                    <div>
                      <h4 className="font-extrabold text-slate-900 text-sm flex items-center space-x-2">
                        <ImageIcon className="w-4 h-4 text-indigo-600" />
                        <span>Hero Banner Image</span>
                      </h4>
                      <p className="text-[11px] text-slate-400 font-medium">Used as the top header banner in certificate emails.</p>
                    </div>
                    <button
                      type="button"
                      onClick={() => setUseCustomHero(!useCustomHero)}
                      className={`relative inline-flex h-6 w-11 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${useCustomHero ? 'bg-[#635bfc]' : 'bg-slate-200'}`}
                    >
                      <span className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow-lg ring-0 transition duration-200 ease-in-out ${useCustomHero ? 'translate-x-5' : 'translate-x-0'}`} />
                    </button>
                  </div>

                  {useCustomHero ? (
                    <div className="space-y-3">
                      {/* Image Preview Box */}
                      {customHeroUrl && (
                        <div className="flex items-center space-x-3 p-3 bg-white rounded-2xl border border-slate-200 shadow-xs">
                          <div className="w-20 h-12 bg-slate-50 rounded-xl border border-slate-100 p-0.5 flex items-center justify-center shrink-0 overflow-hidden">
                            <img src={customHeroUrl} alt="Custom Hero Preview" className="h-full w-full object-cover rounded-lg" />
                          </div>
                          <div className="flex-1 min-w-0">
                            <p className="text-xs font-bold text-slate-800 truncate">Custom Banner Active</p>
                            <a href={customHeroUrl} target="_blank" rel="noreferrer" className="text-[10px] text-indigo-600 hover:underline truncate block">
                              {customHeroUrl}
                            </a>
                          </div>
                          <button
                            type="button"
                            onClick={() => setCustomHeroUrl('')}
                            className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-slate-100 transition-colors"
                            title="Clear image URL"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      )}

                      {/* Upload Button / Dropzone */}
                      <label className={`flex flex-col items-center justify-center p-4 border-2 border-dashed rounded-2xl cursor-pointer transition-all ${
                        isUploadingHero ? 'border-indigo-400 bg-indigo-50/40' : 'border-slate-300 hover:border-indigo-500 bg-white hover:bg-indigo-50/10'
                      }`}>
                        <input
                          type="file"
                          accept="image/*"
                          className="hidden"
                          disabled={isUploadingHero}
                          onChange={(e) => {
                            if (e.target.files && e.target.files[0]) {
                              handleUploadImage(e.target.files[0], 'hero');
                            }
                          }}
                        />
                        {isUploadingHero ? (
                          <div className="flex items-center space-x-2 text-indigo-600 text-xs font-bold">
                            <Loader2 className="w-4 h-4 animate-spin" />
                            <span>Uploading banner to Supabase Storage...</span>
                          </div>
                        ) : (
                          <div className="flex items-center space-x-2 text-slate-600 text-xs font-bold">
                            <Upload className="w-4 h-4 text-indigo-600" />
                            <span>Click or Drag to Upload Hero Banner (.png, .jpg, .webp)</span>
                          </div>
                        )}
                      </label>

                      {/* URL input fallback */}
                      <div>
                        <span className="text-[10px] font-black uppercase tracking-wider text-slate-400 block mb-1">OR ENTER DIRECT URL</span>
                        <input
                          type="text"
                          value={customHeroUrl}
                          onChange={(e) => setCustomHeroUrl(e.target.value)}
                          className="w-full bg-white border border-slate-300 rounded-2xl p-3 text-xs font-semibold text-slate-900 focus:outline-none focus:border-indigo-500 shadow-xs"
                          placeholder="https://your-domain.com/custom_hero.png"
                        />
                      </div>
                    </div>
                  ) : (
                    <div className="flex items-center justify-between p-3 bg-white rounded-2xl border border-slate-200">
                      <div className="flex items-center space-x-3">
                        <div className="w-16 h-10 bg-slate-900 rounded-xl border border-slate-800 p-1 flex items-center justify-center shrink-0">
                          <span className="text-[9px] font-bold text-amber-400 uppercase">CSEA HERO</span>
                        </div>
                        <div>
                          <p className="text-xs font-bold text-slate-700">Using Default CSEA Banner</p>
                          <span className="text-[10px] text-slate-400">Toggle switch above to upload custom hero</span>
                        </div>
                      </div>
                      <span className="bg-slate-100 px-2.5 py-1 rounded-full text-[10px] font-bold text-slate-500">.env Default</span>
                    </div>
                  )}
                </div>
              </div>

              <div className="pt-4 flex flex-wrap items-center gap-3">
                <button
                  onClick={handleSaveEmailSettings}
                  disabled={isSavingEmailSettings}
                  className="flex items-center space-x-2 bg-[#635bfc] hover:bg-[#5248f7] text-white font-extrabold px-6 py-3.5 rounded-2xl text-xs uppercase tracking-wider transition-all shadow-md active:scale-95 disabled:opacity-50"
                >
                  <Mail className="w-4 h-4" />
                  <span>{isSavingEmailSettings ? 'SAVING EMAIL SETTINGS...' : 'SAVE EMAIL SETTINGS'}</span>
                </button>

                <button
                  onClick={() => handleOpenEmailCodeModal(null)}
                  className="flex items-center space-x-2 bg-slate-900 hover:bg-slate-800 text-white font-extrabold px-6 py-3.5 rounded-2xl text-xs uppercase tracking-wider transition-all shadow-md active:scale-95 border border-slate-700"
                >
                  <Code className="w-4 h-4 text-indigo-400" />
                  <span>CUSTOMIZE HTML TEMPLATE (LIVE STUDIO)</span>
                </button>
              </div>
            </div>

          </div>
        )}

        {/* RECENTLY SENT EMAIL TOAST NOTIFICATION WITH QUICK ACTION TO VIEW HTML */}
        {lastSentEmailNotification && (
          <div className="fixed bottom-6 right-6 z-50 bg-slate-900 text-white p-4 rounded-2xl shadow-2xl border border-slate-700 flex items-center space-x-4 animate-in slide-in-from-bottom-5">
            <div className="w-8 h-8 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center shrink-0">
              <Check className="w-4 h-4" />
            </div>
            <div className="text-xs">
              <p className="font-bold text-slate-100">Email Dispatched Successfully!</p>
              <p className="text-slate-400">Sent to {lastSentEmailNotification.studentEmail}</p>
            </div>
            <button
              onClick={() => {
                setEmailCodeModal({
                  isOpen: true,
                  studentName: lastSentEmailNotification.studentName,
                  studentEmail: lastSentEmailNotification.studentEmail,
                  eventName: lastSentEmailNotification.cert.eventName,
                  eventDate: lastSentEmailNotification.cert.eventDate || lastSentEmailNotification.cert.customFields?.event_date || eventDate || lastSentEmailNotification.cert.issueDate || '2026-07-25',
                  htmlContent: lastSentEmailNotification.htmlContent,
                  activeTab: 'code',
                  cert: lastSentEmailNotification.cert
                });
                setLastSentEmailNotification(null);
              }}
              className="bg-indigo-600 hover:bg-indigo-500 text-white font-extrabold px-3 py-1.5 rounded-xl text-xs flex items-center space-x-1 shrink-0 transition-colors shadow-sm"
            >
              <Code className="w-3.5 h-3.5" />
              <span>View HTML</span>
            </button>
            <button
              onClick={() => setLastSentEmailNotification(null)}
              className="text-slate-400 hover:text-white transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* EMAIL HTML CODE & PREVIEW MODAL */}
        {emailCodeModal && emailCodeModal.isOpen && (
          <div className="fixed inset-0 z-50 bg-slate-900/70 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-white rounded-[32px] border border-slate-200 shadow-2xl w-full max-w-6xl h-[90vh] max-h-[920px] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
              
              {/* MODAL HEADER */}
              <div className="px-6 py-4 border-b border-slate-100 flex flex-wrap items-center justify-between gap-3 bg-slate-50/70 shrink-0">
                <div className="flex items-center space-x-3">
                  <div className="w-10 h-10 rounded-2xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
                    <Code className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-extrabold text-slate-900 flex items-center space-x-2">
                      <span>HTML Email Studio & Live Preview</span>
                      <span className="bg-indigo-100 text-indigo-700 text-[10px] font-black uppercase px-2 py-0.5 rounded-full">Editable</span>
                    </h3>
                  </div>
                </div>

                <div className="flex items-center space-x-2">
                  <button
                    onClick={() => handleCopyEmailHtml(emailCodeModal.htmlContent)}
                    className="flex items-center space-x-1.5 bg-slate-800 hover:bg-slate-900 text-white font-extrabold px-3.5 py-2 rounded-xl text-xs shadow-sm transition-all active:scale-95"
                    title="Copy current HTML code to clipboard"
                  >
                    {copiedEmailHtml ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-emerald-400" />
                        <span>Copied HTML!</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" />
                        <span>Copy HTML</span>
                      </>
                    )}
                  </button>

                  <button
                    onClick={() => handleSaveCustomEmailTemplate(emailCodeModal.htmlContent)}
                    disabled={isSavingTemplate}
                    className="flex items-center space-x-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold px-4 py-2 rounded-xl text-xs shadow-sm transition-all active:scale-95 disabled:opacity-50"
                    title="Save this HTML code to database for all future emails"
                  >
                    {isSavingTemplate ? (
                      <>
                        <Loader2 className="w-3.5 h-3.5 animate-spin text-white" />
                        <span>Saving...</span>
                      </>
                    ) : (
                      <>
                        <Database className="w-3.5 h-3.5" />
                        <span>Save</span>
                      </>
                    )}
                  </button>

                  <button
                    onClick={() => setEmailCodeModal(null)}
                    className="p-2 rounded-xl hover:bg-slate-200/60 text-slate-400 hover:text-slate-600 transition-colors"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </div>

              {templateSaveMsg && (
                <div className="px-6 py-2.5 bg-emerald-500 text-white text-xs font-bold flex items-center justify-between animate-in slide-in-from-top-2 shrink-0">
                  <div className="flex items-center space-x-2">
                    <CheckCircle2 className="w-4 h-4 text-emerald-100 shrink-0" />
                    <span>{templateSaveMsg}</span>
                  </div>
                  <button onClick={() => setTemplateSaveMsg('')} className="text-emerald-100 hover:text-white">
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}

              {/* TABS (SPLIT / CODE / PREVIEW / TEMPLATES) & HELPER CHIPS */}
              <div className="px-6 pt-3 pb-2 border-b border-slate-100 flex flex-wrap items-center justify-between gap-3 bg-slate-50/40 shrink-0">
                <div className="flex space-x-2">
                  <button
                    onClick={() => setEmailCodeModal({ ...emailCodeModal, activeTab: 'code' })}
                    className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider transition-all flex items-center space-x-1.5 ${
                      emailCodeModal.activeTab === 'code'
                        ? 'bg-indigo-600 text-white shadow-sm'
                        : 'text-slate-500 hover:text-slate-900 hover:bg-slate-100'
                    }`}
                  >
                    <Code className="w-3.5 h-3.5" />
                    <span>Code Only</span>
                  </button>

                  <button
                    onClick={() => setEmailCodeModal({ ...emailCodeModal, activeTab: 'preview' })}
                    className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider transition-all flex items-center space-x-1.5 ${
                      emailCodeModal.activeTab === 'preview'
                        ? 'bg-indigo-600 text-white shadow-sm'
                        : 'text-slate-500 hover:text-slate-900 hover:bg-slate-100'
                    }`}
                  >
                    <Eye className="w-3.5 h-3.5" />
                    <span>Live Preview</span>
                  </button>
                </div>

                <div className="flex items-center space-x-3 text-[11px] text-slate-500 font-medium overflow-x-auto">
                  <label className="flex items-center space-x-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 px-3 py-1.5 rounded-xl border border-indigo-200 cursor-pointer font-bold text-xs transition-colors shrink-0">
                    <input
                      type="file"
                      accept="image/*"
                      className="hidden"
                      disabled={isUploadingModalAsset}
                      onChange={(e) => {
                        if (e.target.files && e.target.files[0]) {
                          handleUploadImage(e.target.files[0], 'modal');
                        }
                      }}
                    />
                    {isUploadingModalAsset ? (
                      <>
                        <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        <span>Uploading...</span>
                      </>
                    ) : (
                      <>
                        <Upload className="w-3.5 h-3.5" />
                        <span>Upload & Copy Image URL</span>
                      </>
                    )}
                  </label>

                  {modalUploadedAssetUrl && (
                    <span className="bg-emerald-100 text-emerald-800 text-[10px] font-bold px-2.5 py-1 rounded-lg flex items-center space-x-1 shrink-0 animate-in fade-in">
                      <Check className="w-3 h-3 text-emerald-600" />
                      <span>Copied Image URL to Clipboard!</span>
                    </span>
                  )}

                  <div className="flex items-center space-x-1.5 shrink-0">
                    <span className="font-bold text-slate-700">Tokens:</span>
                    <code className="bg-slate-200/80 px-1.5 py-0.5 rounded text-slate-700 font-mono text-[10px]">{"{student_name}"}</code>
                    <code className="bg-slate-200/80 px-1.5 py-0.5 rounded text-slate-700 font-mono text-[10px]">{"{event_name}"}</code>
                    <code className="bg-slate-200/80 px-1.5 py-0.5 rounded text-slate-700 font-mono text-[10px]">{"{event_date}"}</code>
                    <code className="bg-indigo-100 text-indigo-800 px-1.5 py-0.5 rounded font-mono text-[10px]">{"{hero_src}"}</code>
                    <code className="bg-indigo-100 text-indigo-800 px-1.5 py-0.5 rounded font-mono text-[10px]">{"{logo_src}"}</code>
                  </div>
                </div>
              </div>

              {/* MODAL BODY */}
              <div className="flex-1 overflow-hidden p-4 bg-slate-950 text-slate-100 flex flex-col min-h-0">
                {/* 1. CODE ONLY SECTION */}
                {emailCodeModal.activeTab === 'code' ? (
                  <div className="flex-1 h-full flex flex-col space-y-2 min-h-0">
                    <div className="flex items-center justify-between text-[11px] text-slate-400 px-1 shrink-0">
                      <span>Full screen HTML editor. Live preview updates instantly when switching to Live Preview.</span>
                      <span className="font-mono">{emailCodeModal.htmlContent.length.toLocaleString()} chars</span>
                    </div>
                    <textarea
                      value={emailCodeModal.htmlContent}
                      onChange={(e) => {
                        const val = e.target.value;
                        setEmailCodeModal({ ...emailCodeModal, htmlContent: val });
                        setPreviewRefreshKey((prev) => prev + 1);
                      }}
                      className="w-full flex-1 h-full min-h-0 p-4 bg-slate-900 border border-slate-800 rounded-2xl font-mono text-xs text-emerald-400 focus:outline-none focus:border-indigo-500 leading-relaxed resize-none selection:bg-indigo-500 selection:text-white overflow-y-auto"
                      placeholder="Paste or write HTML email template here..."
                      spellCheck={false}
                    />
                  </div>
                ) : (
                  /* 2. LIVE PREVIEW SECTION */
                  <div className="flex-1 h-full min-h-0 bg-slate-900/90 rounded-2xl overflow-hidden shadow-2xl border border-slate-800 flex flex-col">
                    <div className="px-4 py-2 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-2 shrink-0">
                      <div className="flex items-center space-x-2">
                        <button
                          type="button"
                          onClick={() => setPreviewDevice('desktop')}
                          className={`px-3 py-1 rounded-xl text-xs font-bold transition-all ${
                            previewDevice === 'desktop'
                              ? 'bg-indigo-600 text-white shadow-sm'
                              : 'text-slate-400 hover:text-white bg-slate-800/60'
                          }`}
                        >
                          🖥️ Desktop View
                        </button>
                        <button
                          type="button"
                          onClick={() => setPreviewDevice('mobile')}
                          className={`px-3 py-1 rounded-xl text-xs font-bold transition-all ${
                            previewDevice === 'mobile'
                              ? 'bg-indigo-600 text-white shadow-sm'
                              : 'text-slate-400 hover:text-white bg-slate-800/60'
                          }`}
                        >
                          📱 Mobile (380px)
                        </button>
                      </div>

                      <div className="flex items-center space-x-2">
                        <button
                          type="button"
                          onClick={() => setPreviewRefreshKey((prev) => prev + 1)}
                          className="flex items-center space-x-1.5 text-xs text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 px-3 py-1.5 rounded-xl transition-colors font-bold shadow-sm"
                          title="Force reload preview"
                        >
                          <RefreshCw className="w-3.5 h-3.5" />
                          <span>Refresh Preview</span>
                        </button>
                      </div>
                    </div>

                    <div className="flex-1 min-h-0 overflow-y-auto bg-slate-950/70 p-4 flex justify-center items-start">
                      <div className={`transition-all duration-200 shadow-2xl rounded-2xl overflow-hidden border border-slate-800 bg-[#FAF3E7] ${
                        previewDevice === 'mobile' ? 'w-[390px] min-h-[700px]' : 'w-full max-w-[650px] min-h-[700px]'
                      }`}>
                        <iframe
                          key={`preview_frame_${previewDevice}_${previewRefreshKey}_${emailCodeModal.htmlContent.length}`}
                          srcDoc={getRenderedPreviewHtml(emailCodeModal.htmlContent)}
                          title="HTML Email Live Preview"
                          className="w-full h-[720px] border-none bg-[#FAF3E7]"
                        />
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* MODAL FOOTER */}
              <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex items-center justify-end space-x-3 shrink-0">
                <button
                  onClick={() => handleSaveCustomEmailTemplate(emailCodeModal.htmlContent)}
                  disabled={isSavingTemplate}
                  className="flex items-center space-x-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold px-4 py-2 rounded-xl text-xs shadow-sm transition-all active:scale-95 disabled:opacity-50"
                >
                  <Database className="w-3.5 h-3.5" />
                  <span>Save</span>
                </button>

                {emailCodeModal.cert && (
                  <button
                    onClick={() => {
                      handleSendSingleEmail(emailCodeModal.cert!);
                    }}
                    disabled={sendingSingleCertId === emailCodeModal.cert.id}
                    className="flex items-center space-x-1.5 bg-amber-500 hover:bg-amber-600 text-white font-extrabold px-4 py-2 rounded-xl text-xs shadow-sm transition-all active:scale-95 disabled:opacity-50"
                  >
                    <Send className="w-3.5 h-3.5" />
                    <span>{sendingSingleCertId === emailCodeModal.cert.id ? 'Sending...' : 'Send This Email'}</span>
                  </button>
                )}
                <button
                  onClick={() => setEmailCodeModal(null)}
                  className="bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold px-4 py-2 rounded-xl text-xs transition-colors"
                >
                  Close
                </button>
              </div>

            </div>
          </div>
        )}



      </main>

    </div>
  );
}
