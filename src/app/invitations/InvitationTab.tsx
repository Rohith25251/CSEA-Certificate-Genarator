'use client';

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '@/lib/supabase';
import { 
  Upload, 
  Mail, 
  Database, 
  Eye, 
  FileText, 
  AlertCircle, 
  CheckCircle2, 
  ArrowLeft, 
  ArrowRight, 
  Play, 
  Search, 
  RefreshCw,
  Code,
  Sparkles,
  ListTodo,
  Maximize2
} from 'lucide-react';

const DEFAULT_TEMPLATE = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    .card {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      max-width: 600px;
      margin: 20px auto;
      background: #ffffff;
      border-radius: 16px;
      overflow: hidden;
      border: 1px solid #e2e8f0;
      box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    .header {
      background: linear-gradient(135deg, #1e3a8a, #0d9488);
      color: #ffffff;
      padding: 30px 20px;
      text-align: center;
    }
    .content {
      padding: 30px;
      color: #334155;
      line-height: 1.6;
    }
    .btn {
      display: inline-block;
      padding: 12px 24px;
      background: #0284c7;
      color: #ffffff !important;
      text-decoration: none;
      border-radius: 8px;
      font-weight: bold;
      margin-top: 20px;
      text-align: center;
    }
    .footer {
      background: #f8fafc;
      padding: 20px;
      text-align: center;
      font-size: 11px;
      color: #64748b;
      border-top: 1px solid #f1f5f9;
    }
  </style>
</head>
<body style="background-color: #f1f5f9; padding: 10px; margin: 0;">
  <div class="card">
    <div class="header">
      <h1 style="margin: 0; font-size: 24px; font-weight: 800; letter-spacing: -0.5px;">CSE Association</h1>
      <p style="margin: 5px 0 0; font-size: 13px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px;">Kongu Engineering College</p>
    </div>
    <div class="content">
      <h2 style="color: #0f172a; margin-top: 0;">Hello, {{Name}}! 👋</h2>
      <p>We are delighted to invite you to our upcoming CSEA technical session: <strong style="color: #0369a1;">{{EventName}}</strong>.</p>
      <p>This exclusive workshop is designed to equip you with critical skills, hands-on labs, and real-world project demonstrations guided by experienced experts.</p>
      
      <div style="background: #f0fdf4; border-left: 4px solid #16a34a; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <p style="margin: 0; font-size: 14px; color: #14532d;"><strong>📅 Date:</strong> {{Date}}</p>
        <p style="margin: 5px 0 0; font-size: 14px; color: #14532d;"><strong>📍 Venue:</strong> CSE Seminar Hall (Main Block)</p>
      </div>

      <p>Click the link below to confirm your attendance and secure your seat. We look forward to seeing you there!</p>
      
      <div style="text-align: center;">
        <a href="https://csea.kgec.edu/rsvp" class="btn">RSVP / Confirm Seat</a>
      </div>
    </div>
    <div class="footer">
      <p>You received this email because you are an active member of CSEA Kongu.</p>
      <p>© 2026 CSEA - Department of Computer Science and Engineering. All rights reserved.</p>
    </div>
  </div>
</body>
</html>`;

interface InvitationLog {
  id?: string;
  recipient_name: string;
  recipient_email: string;
  event_name: string;
  subject: string;
  status: string;
  error_message?: string;
  sent_at: string;
  custom_data?: any;
}

function parseAndFormatDate(val: any): string {
  if (val === null || val === undefined) return '';
  
  const pad = (n: number) => String(n).padStart(2, '0');
  
  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      const day = val.getUTCDate();
      const month = val.getUTCMonth() + 1;
      const year = val.getUTCFullYear();
      return `${pad(day)}-${pad(month)}-${year}`;
    }
    return '';
  }

  const valStr = String(val).trim();
  if (!valStr) return '';

  // Try to match DD-MM-YYYY or MM-DD-YYYY
  const dmyRegex = /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/;
  const dmyMatch = valStr.match(dmyRegex);
  if (dmyMatch) {
    const p1 = parseInt(dmyMatch[1], 10);
    const p2 = parseInt(dmyMatch[2], 10);
    const year = parseInt(dmyMatch[3], 10);
    if (p1 > 12) {
      return `${pad(p1)}-${pad(p2)}-${year}`;
    }
    if (p2 > 12) {
      return `${pad(p2)}-${pad(p1)}-${year}`;
    }
    return `${pad(p1)}-${pad(p2)}-${year}`;
  }

  // Try to match M/D/YY or MM/DD/YY
  const mdyShortRegex = /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2})$/;
  const mdyShortMatch = valStr.match(mdyShortRegex);
  if (mdyShortMatch) {
    const p1 = parseInt(mdyShortMatch[1], 10);
    const p2 = parseInt(mdyShortMatch[2], 10);
    let year = parseInt(mdyShortMatch[3], 10);
    year = year < 50 ? 2000 + year : 1900 + year;
    if (p1 > 12) {
      return `${pad(p1)}-${pad(p2)}-${year}`;
    }
    if (p2 > 12) {
      return `${pad(p2)}-${pad(p1)}-${year}`;
    }
    return `${pad(p2)}-${pad(p1)}-${year}`;
  }

  // Try to match YYYY-MM-DD or YYYY/MM/DD
  const ymdRegex = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/;
  const ymdMatch = valStr.match(ymdRegex);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10);
    const day = parseInt(ymdMatch[3], 10);
    return `${pad(day)}-${pad(month)}-${year}`;
  }

  // Check if it's a number (Excel serial)
  const num = Number(valStr);
  if (!isNaN(num) && num > 10000 && num < 100000) {
    const days = num - 25569;
    const date = new Date(Math.round(days * 86400 * 1000));
    if (!isNaN(date.getTime())) {
      const day = date.getUTCDate();
      const month = date.getUTCMonth() + 1;
      const year = date.getUTCFullYear();
      return `${pad(day)}-${pad(month)}-${year}`;
    }
  }

  return valStr;
}

const getTodayFormatted = () => {
  const today = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${pad(today.getDate())}-${pad(today.getMonth() + 1)}-${today.getFullYear()}`;
};

const PYTHON_API_URL = process.env.NEXT_PUBLIC_API_URL || 'http://localhost:8000';

export default function InvitationTab() {
  const [stage, setStage] = useState<number>(1);
  
  // Stage 1: Excel Upload
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [rows, setRows] = useState<any[]>([]);
  const [isExcelParsing, setIsExcelParsing] = useState<boolean>(false);
  const [searchQuery, setSearchQuery] = useState<string>('');

  // Stage 2: Template configuration
  const [subject, setSubject] = useState<string>('Invitation to CSEA Technical Event');
  const [eventName, setEventName] = useState<string>('Data Engineering Mastery');
  const [htmlTemplate, setHtmlTemplate] = useState<string>(DEFAULT_TEMPLATE);
  const [detectedPlaceholders, setDetectedPlaceholders] = useState<string[]>([]);

  // Stage 3: Preview
  const [previewRowIndex, setPreviewRowIndex] = useState<number>(0);
  const [previewHtml, setPreviewHtml] = useState<string>('');
  const [copiedPlaceholder, setCopiedPlaceholder] = useState<string | null>(null);
  
  // History Expansion & Resend States
  const [expandedHistoryId, setExpandedHistoryId] = useState<string | null>(null);
  const [resendingId, setResendingId] = useState<string | null>(null);
  const [isHistoryModalOpen, setIsHistoryModalOpen] = useState<boolean>(false);
  
  // Stage 4: Dispatch
  const [isSending, setIsSending] = useState<boolean>(false);
  const [sendingProgress, setSendingProgress] = useState<{ current: number; total: number; sent: number; failed: number }>({
    current: 0,
    total: 0,
    sent: 0,
    failed: 0
  });
  const [dispatchLogs, setDispatchLogs] = useState<Array<{ name: string; email: string; status: 'sent' | 'failed'; error?: string }>>([]);

  // DB Sync History
  const [history, setHistory] = useState<InvitationLog[]>([]);
  const [historySearch, setHistorySearch] = useState<string>('');
  const [isLoadingHistory, setIsLoadingHistory] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Load history on mount
  useEffect(() => {
    fetchHistory();
  }, []);

  // Update detected placeholders when HTML template changes
  useEffect(() => {
    const doubleBraceMatches = htmlTemplate.match(/\{\{\s*([^}]+?)\s*\}\}/g) || [];
    const doubleAngleMatches = htmlTemplate.match(/<<\s*([^>]+?)\s*>>/g) || [];
    
    const placeholders = Array.from(
      new Set([
        ...doubleBraceMatches.map(m => m.replace(/\{\{|\}\}/g, '').trim()),
        ...doubleAngleMatches.map(m => m.replace(/<<|>>/g, '').trim())
      ])
    );
    setDetectedPlaceholders(placeholders);
  }, [htmlTemplate]);

  // Update live preview when HTML template, row index, or rows change
  useEffect(() => {
    if (rows.length > 0 && htmlTemplate) {
      const activeRow = rows[previewRowIndex] || {};
      
      // Resolve dynamic recipientName and recipientEmail
      let recipientName = 'Guest';
      let recipientEmail = '';

      const nameKeys = ['name', 'recipient name', 'recipient_name', 'student name', 'student_name', 'full name', 'fullname', 'receiver name', 'receiver_name'];
      const emailKeys = ['email', 'email address', 'email_address', 'recipient email', 'recipient_email', 'student email', 'student_email', 'mail', 'receiver email', 'receiver_email'];

      Object.keys(activeRow).forEach(k => {
        const kLower = k.toLowerCase().trim();
        if (nameKeys.includes(kLower)) recipientName = String(activeRow[k]).trim();
        if (emailKeys.includes(kLower)) recipientEmail = String(activeRow[k]).trim();
      });

      if (!recipientEmail) {
        const mailKey = Object.keys(activeRow).find(k => k.toLowerCase().includes('mail'));
        if (mailKey) recipientEmail = String(activeRow[mailKey]).trim();
      }
      if (recipientName === 'Guest') {
        const nameKey = Object.keys(activeRow).find(k => k.toLowerCase().includes('name'));
        if (nameKey) recipientName = String(activeRow[nameKey]).trim();
      }

      // Try to find if Excel already has an EventName or Date value
      let excelEventName = '';
      let excelDate = '';
      
      Object.keys(activeRow).forEach(k => {
        const kLower = k.toLowerCase().trim();
        if (kLower === 'eventname' || kLower === 'event_name' || kLower === 'event') {
          excelEventName = String(activeRow[k]).trim();
        }
        if (kLower === 'date') {
          excelDate = String(activeRow[k]).trim();
        }
      });

      const finalEventName = excelEventName || eventName;
      const finalDate = excelDate || getTodayFormatted();

      // Inject standard event info and resolved Name/Email aliases
      const enrichedRow: Record<string, any> = {
        ...activeRow,
        Name: recipientName,
        name: recipientName,
        recipient_name: recipientName,
        Email: recipientEmail,
        email: recipientEmail,
        recipient_email: recipientEmail,
        EventName: finalEventName,
        Event_Name: finalEventName,
        Event: finalEventName,
        Subject: subject,
        Date: finalDate
      };

      let rendered = htmlTemplate;
      // Perform local search and replace
      Object.keys(enrichedRow).forEach(k => {
        const val = String(enrichedRow[k] ?? '');
        const escapedKey = k.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        const pattern = new RegExp(`(<<\\s*${escapedKey}\\s*>>|\\{\\{\\s*${escapedKey}\\s*\\}\\})`, 'gi');
        rendered = rendered.replace(pattern, val);
      });
      setPreviewHtml(rendered);
    }
  }, [htmlTemplate, previewRowIndex, rows, eventName, subject]);

  const fetchHistory = async () => {
    setIsLoadingHistory(true);
    try {
      const { data, error } = await supabase
        .from('invitation')
        .select('*')
        .order('created_at', { ascending: false });

      if (error) {
        // Fallback to local python API
        const res = await fetch(`${PYTHON_API_URL}/api/invitations/history`);
        const apiData = await res.json();
        if (apiData.status === 'success') {
          setHistory(apiData.data);
        }
      } else {
        setHistory(data || []);
      }
    } catch (err) {
      console.error('Failed to load history:', err);
    } finally {
      setIsLoadingHistory(false);
    }
  };

  const handleResend = async (item: InvitationLog) => {
    if (!item.recipient_email) return;
    setResendingId(item.id || null);
    try {
      const payloadRow = {
        ...(item.custom_data || {}),
        Name: item.recipient_name,
        name: item.recipient_name,
        recipient_name: item.recipient_name,
        Email: item.recipient_email,
        email: item.recipient_email,
        recipient_email: item.recipient_email,
        EventName: item.event_name,
        Event_Name: item.event_name,
        Event: item.event_name,
        Subject: item.subject,
        Date: new Date().toLocaleDateString('en-US', { day: 'numeric', month: 'short', year: 'numeric' })
      };

      const response = await fetch(`${PYTHON_API_URL}/api/invitations/send-batch`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          subject: item.subject || 'Invitation',
          event_name: item.event_name || 'Event',
          html_template: htmlTemplate,
          rows: [payloadRow]
        })
      });

      const data = await response.json();
      const runResult = data.results?.[0];

      if (response.ok && runResult && runResult.status === 'sent') {
        alert(`Invitation resent to ${item.recipient_email} successfully!`);
        fetchHistory();
      } else {
        alert(`Resend failed: ${runResult?.error || 'Unknown error'}`);
      }
    } catch (err: any) {
      console.error(err);
      alert(`Error resending invitation: ${err?.message || 'Server connection error'}`);
    } finally {
      setResendingId(null);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setExcelFile(file);
    setIsExcelParsing(true);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const buffer = event.target?.result as ArrayBuffer;
        const workbook = XLSX.read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Raw json with header row
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (json.length > 0) {
          const rawHeaders = json[0] as string[];
          const cleanHeaders = rawHeaders.filter(h => h != null && String(h).trim() !== '');
          setHeaders(cleanHeaders);

          const dataRows = XLSX.utils.sheet_to_json(worksheet, { raw: false }) as any[];
          
          // Format cells (especially Excel dates)
          const formattedRows = dataRows.map(row => {
            const newRow: Record<string, any> = {};
            Object.keys(row).forEach(key => {
              const val = row[key];
              const keyLower = key.toLowerCase().trim();
              if (keyLower.includes('date')) {
                newRow[key] = parseAndFormatDate(val);
              } else {
                newRow[key] = val === null || val === undefined ? '' : String(val).trim();
              }
            });
            return newRow;
          });

          setRows(formattedRows);
          setPreviewRowIndex(0);
        }
      } catch (err) {
        alert('Failed to parse Excel sheet. Ensure the file contains structured rows.');
        console.error(err);
      } finally {
        setIsExcelParsing(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Helper to check if a placeholder has a matching Excel column
  const hasMatchingHeader = (placeholder: string) => {
    const pLower = placeholder.trim().toLowerCase();
    
    // Add default enriched fields that are always available
    const standardFields = ['eventname', 'event_name', 'event', 'subject', 'date'];
    if (standardFields.includes(pLower)) return true;

    // Smart mapping for Name placeholder
    const nameKeywords = ['name', 'recipientname', 'recipient_name', 'studentname', 'student_name', 'fullname', 'full_name', 'receivername', 'receiver_name'];
    if (nameKeywords.includes(pLower)) {
      const hasNameHeader = headers.some(h => {
        const hLower = h.toLowerCase().trim();
        return hLower === 'name' || hLower === 'student name' || hLower === 'full name' || hLower.includes('name');
      });
      if (hasNameHeader) return true;
    }
    
    // Smart mapping for Email placeholder
    const emailKeywords = ['email', 'emailaddress', 'email_address', 'recipientemail', 'recipient_email', 'studentemail', 'student_email', 'mail', 'receiveremail', 'receiver_email'];
    if (emailKeywords.includes(pLower)) {
      const hasEmailHeader = headers.some(h => {
        const hLower = h.toLowerCase().trim();
        return hLower === 'email' || hLower === 'email address' || hLower === 'mail' || hLower.includes('mail') || hLower.includes('email');
      });
      if (hasEmailHeader) return true;
    }
    
    return headers.some(h => h.trim().toLowerCase() === pLower);
  };

  const getAvailablePlaceholders = () => {
    const list = [...headers];
    if (!list.some(h => h.toLowerCase() === 'eventname')) list.push('EventName');
    if (!list.some(h => h.toLowerCase() === 'date')) list.push('Date');
    return list;
  };

  const isPlaceholderUsed = (placeholder: string) => {
    const pLower = placeholder.trim().toLowerCase();
    if (detectedPlaceholders.some(dp => dp.toLowerCase() === pLower)) return true;

    const nameKeywords = ['name', 'recipientname', 'recipient_name', 'studentname', 'student_name', 'fullname', 'full_name', 'receivername', 'receiver_name'];
    if (pLower === 'name' || pLower === 'student name' || pLower === 'full name' || pLower.includes('name')) {
      if (detectedPlaceholders.some(dp => nameKeywords.includes(dp.toLowerCase()))) return true;
    }

    const emailKeywords = ['email', 'emailaddress', 'email_address', 'recipientemail', 'recipient_email', 'studentemail', 'student_email', 'mail', 'receiveremail', 'receiver_email'];
    if (pLower === 'email' || pLower === 'email address' || pLower === 'mail' || pLower.includes('mail') || pLower.includes('email')) {
      if (detectedPlaceholders.some(dp => emailKeywords.includes(dp.toLowerCase()))) return true;
    }

    return false;
  };

  const handleCopyPlaceholder = (placeholder: string) => {
    const text = `{{${placeholder}}}`;
    navigator.clipboard.writeText(text);
    setCopiedPlaceholder(placeholder);
    setTimeout(() => setCopiedPlaceholder(null), 2000);
  };

  const executeBatchSend = async () => {
    if (rows.length === 0) {
      alert('Please upload receiver details first.');
      return;
    }

    setIsSending(true);
    setDispatchLogs([]);
    setSendingProgress({ current: 0, total: rows.length, sent: 0, failed: 0 });

    let sentCount = 0;
    let failedCount = 0;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      setSendingProgress(prev => ({ ...prev, current: i + 1 }));

      // Find recipient details in row
      let recipientName = 'Guest';
      let recipientEmail = '';

      const nameKeys = ['name', 'recipient name', 'recipient_name', 'student name', 'student_name', 'full name', 'fullname', 'receiver name', 'receiver_name'];
      const emailKeys = ['email', 'email address', 'email_address', 'recipient email', 'recipient_email', 'student email', 'student_email', 'mail', 'receiver email', 'receiver_email'];

      Object.keys(row).forEach(k => {
        const kLower = k.toLowerCase().trim();
        if (nameKeys.includes(kLower)) recipientName = String(row[k]).trim();
        if (emailKeys.includes(kLower)) recipientEmail = String(row[k]).trim();
      });

      // Substring searches as fallback
      if (!recipientEmail) {
        const mailKey = Object.keys(row).find(k => k.toLowerCase().includes('mail'));
        if (mailKey) recipientEmail = String(row[mailKey]).trim();
      }
      if (recipientName === 'Guest') {
        const nameKey = Object.keys(row).find(k => k.toLowerCase().includes('name'));
        if (nameKey) recipientName = String(row[nameKey]).trim();
      }

      if (!recipientEmail || !recipientEmail.includes('@')) {
        failedCount++;
        setDispatchLogs(prev => [
          ...prev,
          { name: recipientName, email: recipientEmail || 'Missing', status: 'failed', error: 'Missing or invalid email' }
        ]);
        setSendingProgress(prev => ({ ...prev, failed: failedCount }));
        continue;
      }

      // Try to find if Excel already has an EventName or Date value
      let excelEventName = '';
      let excelDate = '';
      
      Object.keys(row).forEach(k => {
        const kLower = k.toLowerCase().trim();
        if (kLower === 'eventname' || kLower === 'event_name' || kLower === 'event') {
          excelEventName = String(row[k]).trim();
        }
        if (kLower === 'date') {
          excelDate = String(row[k]).trim();
        }
      });

      const finalEventName = excelEventName || eventName;
      const finalDate = excelDate || getTodayFormatted();

      // Add enriched variables to custom_data for placeholder replacing
      const payloadRow = {
        ...row,
        EventName: finalEventName,
        Event_Name: finalEventName,
        Event: finalEventName,
        Subject: subject,
        Date: finalDate
      };

      try {
        // We call the server endpoint with a single row to stream progress to client
        const response = await fetch(`${PYTHON_API_URL}/api/invitations/send-batch`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            subject: subject,
            event_name: finalEventName,
            html_template: htmlTemplate,
            rows: [payloadRow]
          })
        });

        const data = await response.json();
        const runResult = data.results?.[0];

        if (response.ok && runResult && runResult.status === 'sent') {
          sentCount++;
          setDispatchLogs(prev => [
            ...prev,
            { name: recipientName, email: recipientEmail, status: 'sent' }
          ]);
          setSendingProgress(prev => ({ ...prev, sent: sentCount }));
        } else {
          failedCount++;
          setDispatchLogs(prev => [
            ...prev,
            { name: recipientName, email: recipientEmail, status: 'failed', error: runResult?.error || 'Failed to dispatch' }
          ]);
          setSendingProgress(prev => ({ ...prev, failed: failedCount }));
        }
      } catch (err: any) {
        failedCount++;
        setDispatchLogs(prev => [
          ...prev,
          { name: recipientName, email: recipientEmail, status: 'failed', error: err?.message || 'Server connection error' }
        ]);
        setSendingProgress(prev => ({ ...prev, failed: failedCount }));
      }

      // Small delay between calls to preserve SMTP sanity
      await new Promise(r => setTimeout(r, 600));
    }

    setIsSending(false);
    fetchHistory();
  };

  const loadSampleTemplate = () => {
    setHtmlTemplate(DEFAULT_TEMPLATE);
  };

  // Filtered rows for Excel preview table
  const filteredRows = rows.filter(r => {
    if (!searchQuery) return true;
    return Object.values(r).some(val => 
      String(val).toLowerCase().includes(searchQuery.toLowerCase())
    );
  });

  // Filtered history rows
  const filteredHistory = history.filter(h => {
    if (!historySearch) return true;
    return (
      h.recipient_name.toLowerCase().includes(historySearch.toLowerCase()) ||
      h.recipient_email.toLowerCase().includes(historySearch.toLowerCase()) ||
      h.event_name.toLowerCase().includes(historySearch.toLowerCase()) ||
      h.subject.toLowerCase().includes(historySearch.toLowerCase()) ||
      h.status.toLowerCase().includes(historySearch.toLowerCase())
    );
  });

  return (
    <div className="max-w-7xl mx-auto px-4 py-8">
      {/* Title block */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 pb-6 border-b border-slate-200">
        <div>
          <h1 className="text-3xl font-extrabold text-[#082849] flex items-center gap-2">
            <Mail className="w-8 h-8 text-indigo-600 animate-pulse" />
            HTML Invitation Dispatcher
          </h1>
          <p className="text-sm text-slate-500 mt-1">
            Upload receiver details, paste your HTML email body, preview dynamic fields, and batch dispatch invites.
          </p>
        </div>

      </div>

      {/* Stepper Wizard Indicator */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm mb-8 flex justify-between items-center max-w-3xl mx-auto">
        {[
          { num: 1, title: 'Upload Sheet' },
          { num: 2, title: 'HTML Template' },
          { num: 3, title: 'Live Preview' },
          { num: 4, title: 'Send Emails' }
        ].map((s) => (
          <button
            key={s.num}
            onClick={() => rows.length > 0 && setStage(s.num)}
            disabled={rows.length === 0 && s.num > 1}
            className={`flex items-center space-x-2.5 transition-all group ${
              stage === s.num
                ? 'text-indigo-600 font-bold'
                : rows.length > 0
                ? 'text-slate-700 hover:text-indigo-500'
                : 'text-slate-300 cursor-not-allowed'
            }`}
          >
            <span className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-black transition-all ${
              stage === s.num
                ? 'bg-indigo-600 text-white shadow-md shadow-indigo-200'
                : stage > s.num
                ? 'bg-emerald-500 text-white'
                : 'bg-slate-100 text-slate-400 group-hover:bg-slate-200'
            }`}>
              {s.num}
            </span>
            <span className="text-xs hidden sm:inline">{s.title}</span>
          </button>
        ))}
      </div>

      {/* WIZARD CONTAINER */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        
        {/* Main Work Area (Left 2 cols) */}
        <div className="lg:col-span-2 space-y-6">
          
          {/* STAGE 01: UPLOAD EXCEL */}
          {stage === 1 && (
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-6">
              <div className="flex items-center justify-between">
                <h3 className="font-extrabold text-base text-slate-800 flex items-center gap-2">
                  <Database className="w-5 h-5 text-indigo-500" />
                  Step 1: Upload Receiver Excel Details
                </h3>
                {rows.length > 0 && (
                  <span className="bg-emerald-100 text-emerald-800 text-[10px] font-extrabold uppercase px-2.5 py-1 rounded-full">
                    {rows.length} Rows Loaded
                  </span>
                )}
              </div>

              {/* Form Config */}
              <div className="pb-4 border-b border-slate-100">
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider block text-slate-600">Email Subject Line</label>
                  <input
                    type="text"
                    value={subject}
                    onChange={(e) => setSubject(e.target.value)}
                    placeholder="Enter email subject line..."
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 font-semibold text-slate-800"
                  />
                </div>
              </div>

              {/* Upload Drop Zone */}
              <div 
                onClick={() => fileInputRef.current?.click()}
                className={`border-2 border-dashed rounded-2xl p-8 text-center cursor-pointer transition-all ${
                  excelFile 
                    ? 'border-emerald-300 bg-emerald-50/20 hover:bg-emerald-50/40' 
                    : 'border-slate-300 hover:border-indigo-500 bg-slate-50/40 hover:bg-slate-50'
                }`}
              >
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  accept=".xlsx, .xls, .csv" 
                  onChange={handleFileUpload} 
                  className="hidden" 
                />
                <div className="max-w-md mx-auto space-y-3">
                  <div className="w-12 h-12 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center mx-auto shadow-inner">
                    <Upload className="w-6 h-6" />
                  </div>
                  {excelFile ? (
                    <div>
                      <p className="text-sm font-bold text-slate-800">{excelFile.name}</p>
                      <p className="text-xs text-slate-400 mt-1">{(excelFile.size / 1024).toFixed(1)} KB • Click to upload a different file</p>
                    </div>
                  ) : (
                    <div>
                      <p className="text-sm font-bold text-slate-700">Drag & Drop receiver sheet here, or browse files</p>
                      <p className="text-xs text-slate-400 mt-1">Supports Excel Workbook (.xlsx, .xls) and CSV files</p>
                    </div>
                  )}
                </div>
              </div>

              {isExcelParsing && (
                <div className="flex items-center justify-center space-x-2 text-indigo-600 font-bold py-6">
                  <RefreshCw className="w-5 h-5 animate-spin" />
                  <span>Parsing Excel Sheet...</span>
                </div>
              )}

              {/* Preview Table */}
              {rows.length > 0 && !isExcelParsing && (
                <div className="space-y-4 pt-4 border-t border-slate-100">
                  <div className="flex items-center justify-between">
                    <h4 className="font-extrabold text-xs text-slate-600 uppercase tracking-wider">Sheet Data Preview</h4>
                    <div className="relative w-64">
                      <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-2.5" />
                      <input
                        type="text"
                        placeholder="Search sheet rows..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="w-full pl-8 pr-3 py-1.5 rounded-xl border border-slate-200 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500"
                      />
                    </div>
                  </div>

                  <div className="border border-slate-150 rounded-xl overflow-hidden max-h-60 overflow-y-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead className="bg-slate-50 sticky top-0 border-b border-slate-200">
                        <tr>
                          {headers.map((h, i) => (
                            <th key={i} className="py-2.5 px-4 font-bold text-slate-600">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 text-slate-700 font-medium">
                        {filteredRows.slice(0, 10).map((row, rIdx) => (
                          <tr key={rIdx} className="hover:bg-slate-50/50">
                            {headers.map((h, cIdx) => (
                              <td key={cIdx} className="py-2 px-4 whitespace-nowrap overflow-hidden max-w-xs text-ellipsis">
                                {String(row[h] ?? '')}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {filteredRows.length > 10 && (
                    <p className="text-[10px] text-slate-400 text-right font-bold">
                      Showing first 10 of {filteredRows.length} total matched rows.
                    </p>
                  )}

                  {/* Actions */}
                  <div className="flex justify-end pt-4">
                    <button
                      onClick={() => setStage(2)}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold text-xs px-5 py-2.5 rounded-xl flex items-center gap-1.5 shadow-md shadow-indigo-150 transition-all hover:scale-102"
                    >
                      Configure HTML Template
                      <ArrowRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* STAGE 02: HTML TEMPLATE & CONFIG */}
          {stage === 2 && (
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-6">
              <h3 className="font-extrabold text-base text-slate-800 flex items-center gap-2">
                <Code className="w-5 h-5 text-indigo-500" />
                Step 2: Paste HTML Email Template
              </h3>

              {/* Form Config */}
              <div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider block">Email Subject</label>
                  <input
                    type="text"
                    value={subject}
                    onChange={(e) => setSubject(e.target.value)}
                    placeholder="Enter email subject line..."
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
              </div>

              {/* Editor Textarea */}
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">HTML Document Code</label>
                  <button
                    onClick={loadSampleTemplate}
                    className="text-xs text-indigo-600 hover:text-indigo-700 font-extrabold flex items-center gap-1.5"
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    Load Sample Template
                  </button>
                </div>
                
                <textarea
                  value={htmlTemplate}
                  onChange={(e) => setHtmlTemplate(e.target.value)}
                  placeholder="Paste your raw HTML template here..."
                  className="w-full h-80 px-4 py-3 font-mono text-xs rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-900 text-slate-200"
                />
              </div>

              {/* Navigation */}
              <div className="flex justify-between items-center pt-4 border-t border-slate-100">
                <button
                  onClick={() => setStage(1)}
                  className="border border-slate-200 hover:bg-slate-50 text-slate-700 font-bold text-xs px-4 py-2.5 rounded-xl flex items-center gap-1.5"
                >
                  <ArrowLeft className="w-4 h-4" />
                  Back
                </button>

                <button
                  onClick={() => setStage(3)}
                  className="bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold text-xs px-5 py-2.5 rounded-xl flex items-center gap-1.5 shadow-md shadow-indigo-150 transition-all hover:scale-102"
                >
                  Generate Preview
                  <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {/* STAGE 03: LIVE PREVIEW */}
          {stage === 3 && (
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-6">
              <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                <h3 className="font-extrabold text-base text-slate-800 flex items-center gap-2">
                  <Eye className="w-5 h-5 text-indigo-500" />
                  Step 3: Preview Rendered Template
                </h3>
              </div>

              {/* Subject Banner */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs flex items-center gap-3">
                <div className="font-bold text-slate-500 uppercase tracking-wider">Subject Line:</div>
                <div className="font-extrabold text-slate-800">{subject}</div>
              </div>

              {/* Dynamic iFrame Render Sandbox */}
              <div className="border border-slate-250 rounded-2xl overflow-hidden bg-white shadow-inner">
                <iframe
                  title="Invitation HTML Preview"
                  srcDoc={previewHtml}
                  className="w-full h-[450px] border-none"
                  sandbox="allow-same-origin allow-scripts"
                />
              </div>

              {/* Navigation */}
              <div className="flex justify-between items-center pt-4">
                <button
                  onClick={() => setStage(2)}
                  className="border border-slate-200 hover:bg-slate-50 text-slate-700 font-bold text-xs px-4 py-2.5 rounded-xl flex items-center gap-1.5"
                >
                  <ArrowLeft className="w-4 h-4" />
                  Back
                </button>

                <button
                  onClick={() => setStage(4)}
                  className="bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold text-xs px-5 py-2.5 rounded-xl flex items-center gap-1.5 shadow-md shadow-indigo-150 transition-all hover:scale-102"
                >
                  Prepare Dispatcher
                  <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {/* STAGE 04: DISPATCHER PROGRESS */}
          {stage === 4 && (
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-6">
              <h3 className="font-extrabold text-base text-slate-800 flex items-center gap-2">
                <Mail className="w-5 h-5 text-indigo-500" />
                Step 4: Dispatch Invitation Campaign
              </h3>

              <div className="bg-indigo-50/50 border border-indigo-100 rounded-2xl p-6 text-center space-y-4">
                <div className="max-w-md mx-auto space-y-2">
                  <h4 className="font-extrabold text-slate-800 text-sm">You are ready to dispatch emails</h4>
                  <p className="text-xs text-slate-500">
                    We will send emails to <strong>{rows.length} recipients</strong> parsed from the Excel sheet.
                    The invitation details and dispatch status will be logged in the Supabase database.
                  </p>
                </div>

                {!isSending && sendingProgress.current === 0 && (
                  <button
                    onClick={executeBatchSend}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold text-sm px-6 py-3 rounded-xl inline-flex items-center gap-2 shadow-lg shadow-indigo-150 transition-all hover:scale-102"
                  >
                    <Play className="w-4 h-4 fill-white" />
                    Start Campaign Dispatch
                  </button>
                )}
              </div>

              {/* Progress Tracker */}
              {(isSending || sendingProgress.current > 0) && (
                <div className="space-y-4 pt-4">
                  <div className="flex justify-between text-xs font-bold text-slate-600">
                    <span>Sending Queue: {sendingProgress.current} / {sendingProgress.total}</span>
                    <span>{Math.round((sendingProgress.current / sendingProgress.total) * 100)}% Complete</span>
                  </div>

                  {/* Progress bar container */}
                  <div className="w-full bg-slate-100 rounded-full h-3 overflow-hidden">
                    <div 
                      className="bg-indigo-600 h-full rounded-full transition-all duration-300"
                      style={{ width: `${(sendingProgress.current / sendingProgress.total) * 100}%` }}
                    />
                  </div>

                  {/* Stats chips */}
                  <div className="grid grid-cols-3 gap-4 text-center">
                    <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-150">
                      <div className="text-[10px] text-slate-400 font-extrabold uppercase">Queue Total</div>
                      <div className="text-base font-black text-slate-700">{sendingProgress.total}</div>
                    </div>
                    <div className="bg-emerald-50 p-2.5 rounded-xl border border-emerald-100">
                      <div className="text-[10px] text-emerald-600 font-extrabold uppercase">Sent</div>
                      <div className="text-base font-black text-emerald-600">{sendingProgress.sent}</div>
                    </div>
                    <div className="bg-red-50 p-2.5 rounded-xl border border-red-100">
                      <div className="text-[10px] text-red-500 font-extrabold uppercase">Failed</div>
                      <div className="text-base font-black text-red-500">{sendingProgress.failed}</div>
                    </div>
                  </div>

                  {/* Realtime dispatch logs */}
                  <div className="space-y-2">
                    <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider block">Live Delivery Log</h4>
                    <div className="border border-slate-150 rounded-xl max-h-48 overflow-y-auto divide-y divide-slate-100 p-2 bg-slate-50 font-mono text-[10px]">
                      {dispatchLogs.length === 0 ? (
                        <div className="text-center text-slate-400 py-4">Waiting to start dispatch...</div>
                      ) : (
                        dispatchLogs.slice().reverse().map((log, idx) => (
                          <div key={idx} className="py-1.5 px-3 flex items-center justify-between">
                            <span className="text-slate-600 truncate mr-4">
                              [{dispatchLogs.length - idx}] {log.name} &lt;{log.email}&gt;
                            </span>
                            {log.status === 'sent' ? (
                              <span className="text-emerald-600 font-bold flex items-center gap-1">
                                <CheckCircle2 className="w-3 h-3" /> SENT
                              </span>
                            ) : (
                              <span className="text-red-500 font-bold flex items-center gap-1" title={log.error}>
                                <AlertCircle className="w-3 h-3" /> FAILED: {log.error}
                              </span>
                            )}
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </div>
              )}

              {/* Navigation Back */}
              <div className="flex justify-start pt-4 border-t border-slate-100">
                <button
                  onClick={() => setStage(3)}
                  disabled={isSending}
                  className="border border-slate-200 hover:bg-slate-50 text-slate-700 font-bold text-xs px-4 py-2.5 rounded-xl flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <ArrowLeft className="w-4 h-4" />
                  Back to Preview
                </button>
              </div>
            </div>
          )}

        </div>

        {/* SIDE PANEL (Right 1 col): Placeholders and Database Logs */}
        <div className="space-y-6">
          
          {/* PLACEHOLDER METRIC CARD */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <h4 className="font-extrabold text-xs text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <ListTodo className="w-4 h-4 text-indigo-500" />
              Available Placeholders
            </h4>
            <p className="text-slate-400 text-[10px] leading-relaxed">
              Variables derived dynamically from your Excel sheet and system defaults. Click any placeholder pill to copy it to your clipboard.
            </p>

            <div className="space-y-2">
              {getAvailablePlaceholders().map((p, i) => {
                const used = isPlaceholderUsed(p);
                const isCopied = copiedPlaceholder === p;
                return (
                  <div key={i} className="flex items-center justify-between text-xs py-1.5 border-b border-slate-50 last:border-0">
                    <button
                      onClick={() => handleCopyPlaceholder(p)}
                      className={`font-mono px-2 py-0.5 rounded text-[10px] text-left hover:scale-102 transition-all ${
                        isCopied
                          ? 'bg-emerald-500 text-white font-extrabold'
                          : 'bg-slate-100 hover:bg-indigo-50 text-indigo-700 hover:text-indigo-800 font-bold border border-slate-200'
                      }`}
                      title={isCopied ? 'Copied!' : 'Click to copy to clipboard'}
                    >
                      {isCopied ? 'Copied!' : `{{${p}}}`}
                    </button>
                    
                    {used ? (
                      <span className="text-[10px] bg-emerald-50 border border-emerald-100 text-emerald-700 px-2 py-0.5 rounded-full font-bold flex items-center gap-0.5">
                        <CheckCircle2 className="w-2.5 h-2.5" /> Used
                      </span>
                    ) : (
                      <span className="text-[10px] bg-slate-50 border border-slate-200 text-slate-400 px-2 py-0.5 rounded-full font-bold flex items-center gap-0.5">
                        <AlertCircle className="w-2.5 h-2.5" /> Unused
                      </span>
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          {/* HISTORICAL SENT ARCHIVE */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between">
              <h4 className="font-extrabold text-xs text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                <Database className="w-4 h-4 text-indigo-500" />
                Invitation History
              </h4>
              <div className="flex items-center gap-2">
                <button 
                  onClick={fetchHistory}
                  disabled={isLoadingHistory}
                  className="text-slate-400 hover:text-slate-700 transition-colors"
                  title="Refresh logs"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isLoadingHistory ? 'animate-spin' : ''}`} />
                </button>
                <button 
                  onClick={() => setIsHistoryModalOpen(true)}
                  className="text-slate-400 hover:text-slate-700 transition-colors"
                  title="Expand History"
                >
                  <Maximize2 className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>

            <div className="relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-2.5" />
              <input
                type="text"
                placeholder="Search history logs..."
                value={historySearch}
                onChange={(e) => setHistorySearch(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 rounded-xl border border-slate-200 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500"
              />
            </div>

            <div className="space-y-2.5 max-h-80 overflow-y-auto pr-1">
              {isLoadingHistory && history.length === 0 ? (
                <div className="text-center text-xs text-slate-400 py-6">Syncing log directory...</div>
              ) : filteredHistory.length === 0 ? (
                <div className="text-center text-xs text-slate-400 py-6">No historical runs found.</div>
              ) : (
                filteredHistory.map((item) => {
                  const isExpanded = expandedHistoryId === item.id;
                  return (
                    <div 
                      key={item.id} 
                      onClick={() => setExpandedHistoryId(isExpanded ? null : (item.id || null))}
                      className={`p-3 bg-slate-50 hover:bg-indigo-50/20 border rounded-xl text-[10px] space-y-2 transition-all cursor-pointer ${
                        isExpanded ? 'border-indigo-300 shadow-sm ring-1 ring-indigo-100 bg-indigo-50/5' : 'border-slate-150'
                      }`}
                    >
                      {/* Name and Status Header */}
                      <div className="flex justify-between items-start font-bold">
                        <span className="text-slate-800 text-xs truncate mr-3">{item.recipient_name}</span>
                        {item.status === 'sent' ? (
                          <span className="text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded-md text-[9px] font-black uppercase shrink-0">Sent</span>
                        ) : (
                          <span className="text-red-700 bg-red-50 px-1.5 py-0.5 rounded-md text-[9px] font-black uppercase shrink-0">Failed</span>
                        )}
                      </div>
                      
                      {/* Email address */}
                      <div className="text-slate-500 font-semibold">{item.recipient_email}</div>
                      
                      {/* Expanded Section Details */}
                      {isExpanded ? (
                        <div className="space-y-2 pt-2 border-t border-slate-200 mt-2">
                          <div className="flex justify-between text-slate-500 font-medium py-0.5">
                            <span className="font-bold">Event Name:</span>
                            <span className="text-slate-700 text-right">{item.event_name}</span>
                          </div>
                          <div className="flex justify-between text-slate-500 font-medium py-0.5">
                            <span className="font-bold">Subject:</span>
                            <span className="text-slate-700 text-right">{item.subject}</span>
                          </div>
                          
                          {item.custom_data && Object.keys(item.custom_data).map((k) => {
                            const kLower = k.toLowerCase().trim();
                            if (['name', 'email', 'recipientname', 'recipient_name', 'studentname', 'student_name', 'student_email', 'studentemail', 'mail', 'eventname', 'event_name', 'event', 'subject'].includes(kLower)) return null;
                            return (
                              <div key={k} className="flex justify-between text-slate-500 font-medium py-0.5 border-b border-slate-100 last:border-0">
                                <span className="font-bold">{k}:</span>
                                <span className="text-slate-700 text-right">{String(item.custom_data[k])}</span>
                              </div>
                            );
                          })}

                          {/* Error logging */}
                          {item.status === 'failed' && item.error_message && (
                            <div className="bg-red-50 border border-red-100 rounded-lg p-2 text-red-700 font-bold text-[9px]">
                              Error: {item.error_message}
                            </div>
                          )}

                          {/* Resend Action Button */}
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleResend(item);
                            }}
                            disabled={resendingId === item.id}
                            className="w-full mt-2 inline-flex items-center justify-center space-x-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold py-2 px-3 rounded-lg text-[9px] transition-colors disabled:opacity-50 disabled:cursor-not-allowed shadow-sm"
                          >
                            {resendingId === item.id ? (
                              <RefreshCw className="w-3.5 h-3.5 animate-spin text-white" />
                            ) : (
                              <Mail className="w-3.5 h-3.5 text-white" />
                            )}
                            <span>{item.status === 'failed' ? 'Retry Sending' : 'Resend Invitation Email'}</span>
                          </button>
                        </div>
                      ) : (
                        /* Truncated Default view */
                        <div className="flex justify-between items-center text-[9px] text-slate-400 pt-1 border-t border-slate-100">
                          <span className="truncate max-w-[130px] font-semibold">{item.event_name}</span>
                          <span>{new Date(item.sent_at).toLocaleDateString()}</span>
                        </div>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          </div>

        </div>

      </div>

      {/* EXPANDED HISTORY MODAL */}
      {isHistoryModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl border border-slate-200 shadow-2xl max-w-4xl w-full h-[80vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            
            {/* Modal Header */}
            <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-xl bg-indigo-50 text-indigo-600 flex items-center justify-center">
                  <Database className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="font-extrabold text-base text-slate-800">Invitation History Archive</h3>
                  <p className="text-[10px] text-slate-400 font-medium">Browse, search, and resend dispatched invitations</p>
                </div>
              </div>
              
              <div className="flex items-center gap-3">
                <button 
                  onClick={fetchHistory}
                  disabled={isLoadingHistory}
                  className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition-all"
                  title="Refresh logs"
                >
                  <RefreshCw className={`w-4 h-4 ${isLoadingHistory ? 'animate-spin' : ''}`} />
                </button>
                
                <button 
                  onClick={() => setIsHistoryModalOpen(false)}
                  className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition-all font-bold text-xs uppercase text-slate-600"
                  title="Close"
                >
                  Close
                </button>
              </div>
            </div>

            {/* Modal Search */}
            <div className="p-4 border-b border-slate-100 flex items-center gap-4 bg-white">
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
                <input
                  type="text"
                  placeholder="Search by recipient name, email, event or subject..."
                  value={historySearch}
                  onChange={(e) => setHistorySearch(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 rounded-xl border border-slate-200 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500"
                />
              </div>
              {filteredHistory.length > 0 && (
                <span className="text-[10px] bg-slate-100 text-slate-600 font-bold px-2.5 py-1 rounded-full uppercase shrink-0">
                  {filteredHistory.length} Runs Found
                </span>
              )}
            </div>

            {/* Modal Body / Scroll area */}
            <div className="flex-1 overflow-y-auto p-6 bg-slate-50/30">
              {isLoadingHistory && history.length === 0 ? (
                <div className="flex flex-col items-center justify-center py-20 text-slate-400 gap-3">
                  <RefreshCw className="w-8 h-8 animate-spin text-indigo-500" />
                  <span className="text-xs font-bold">Syncing log archive...</span>
                </div>
              ) : filteredHistory.length === 0 ? (
                <div className="text-center text-xs text-slate-400 py-20 font-bold">
                  No matches found for your search queries.
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {filteredHistory.map((item) => {
                    const isExpanded = expandedHistoryId === item.id;
                    return (
                      <div 
                        key={item.id}
                        onClick={() => setExpandedHistoryId(isExpanded ? null : (item.id || null))}
                        className={`p-5 bg-white border rounded-2xl text-xs space-y-3 transition-all cursor-pointer shadow-sm hover:shadow-md ${
                          isExpanded ? 'border-indigo-300 ring-2 ring-indigo-50 bg-indigo-50/5' : 'border-slate-200 hover:border-slate-350'
                        }`}
                      >
                        <div className="flex justify-between items-start font-bold">
                          <span className="text-slate-800 text-sm truncate mr-3">{item.recipient_name}</span>
                          {item.status === 'sent' ? (
                            <span className="text-emerald-700 bg-emerald-50 border border-emerald-100 px-2 py-0.5 rounded-full text-[9px] font-black uppercase shrink-0">Sent</span>
                          ) : (
                            <span className="text-red-700 bg-red-50 border border-red-100 px-2 py-0.5 rounded-full text-[9px] font-black uppercase shrink-0">Failed</span>
                          )}
                        </div>
                        
                        <div className="text-slate-500 font-semibold truncate">{item.recipient_email}</div>
                        
                        <div className="flex justify-between items-center text-[10px] text-slate-400 border-t border-slate-100 pt-2 mt-2">
                          <span className="font-bold truncate max-w-[60%]">{item.event_name}</span>
                          <span className="shrink-0">{new Date(item.sent_at).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}</span>
                        </div>

                        {isExpanded && (
                          <div className="space-y-2 pt-3 border-t border-slate-200 mt-2 animate-in fade-in slide-in-from-top-1 duration-150">
                            <div className="flex justify-between text-[11px] text-slate-500 font-medium py-0.5">
                              <span className="font-bold">Subject:</span>
                              <span className="text-slate-700 text-right">{item.subject}</span>
                            </div>
                            
                            {item.custom_data && Object.keys(item.custom_data).map((k) => {
                              const kLower = k.toLowerCase().trim();
                              if (['name', 'email', 'recipientname', 'recipient_name', 'studentname', 'student_name', 'student_email', 'studentemail', 'mail', 'eventname', 'event_name', 'event', 'subject'].includes(kLower)) return null;
                              return (
                                <div key={k} className="flex justify-between text-[11px] text-slate-500 font-medium py-0.5 border-b border-slate-100 last:border-0">
                                  <span className="font-bold">{k}:</span>
                                  <span className="text-slate-700 text-right">{String(item.custom_data[k])}</span>
                                </div>
                              );
                            })}

                            {item.status === 'failed' && item.error_message && (
                              <div className="bg-red-50 border border-red-100 rounded-xl p-2.5 text-red-700 font-bold text-[10px] mt-2">
                                Error: {item.error_message}
                              </div>
                            )}

                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                handleResend(item);
                              }}
                              disabled={resendingId === item.id}
                              className="w-full mt-3 inline-flex items-center justify-center space-x-2 bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold py-2 px-4 rounded-xl text-xs transition-all disabled:opacity-50 disabled:cursor-not-allowed shadow-md shadow-indigo-100"
                            >
                              {resendingId === item.id ? (
                                <RefreshCw className="w-4 h-4 animate-spin text-white" />
                              ) : (
                                <Mail className="w-4 h-4" />
                              )}
                              <span>Resend Invitation Email</span>
                            </button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

          </div>
        </div>
      )}

    </div>
  );
}
