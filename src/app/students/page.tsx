'use client';

import React, { useState, useEffect } from 'react';
import Link from 'next/link';
import { StoreManager } from '@/lib/store';
import { Student } from '@/lib/db';
import { fetchStudentsFromSupabase } from '@/lib/supabase';
import { Search, Database, Sparkles, RefreshCw } from 'lucide-react';

export default function StudentsPage() {
  const [students, setStudents] = useState<Student[]>([]);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(true);

  const loadStudentData = async () => {
    setIsLoading(true);
    const dbStudents = await fetchStudentsFromSupabase();
    if (dbStudents.length > 0) {
      setStudents(dbStudents);
    } else {
      setStudents(StoreManager.getStudents());
    }
    setIsLoading(false);
  };

  useEffect(() => {
    loadStudentData();
  }, []);

  const filteredStudents = students.filter(s =>
    s.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    s.registerNo.toLowerCase().includes(searchQuery.toLowerCase()) ||
    s.email.toLowerCase().includes(searchQuery.toLowerCase()) ||
    s.section.toLowerCase().includes(searchQuery.toLowerCase())
  );

  return (
    <div className="min-h-screen bg-slate-50 py-10">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        
        {/* Header */}
        <div className="flex flex-col md:flex-row md:items-center md:justify-between mb-8 pb-6 border-b border-slate-200">
          <div>
            <h1 className="text-3xl font-extrabold text-[#082849]">Master Participant Directory</h1>
            <p className="text-sm text-slate-500 mt-1">
              Registered students & participation records synced live from Supabase & Excel uploads
            </p>
          </div>

          <div className="mt-4 md:mt-0 flex items-center space-x-3">
            <button
              onClick={loadStudentData}
              className="inline-flex items-center space-x-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold px-3 py-2 rounded-xl text-xs transition-colors"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
              <span>Refresh Supabase</span>
            </button>

            <div className="bg-blue-100 border border-blue-200 text-blue-800 font-extrabold text-xs px-4 py-2 rounded-xl">
              Total Records: {filteredStudents.length} Students
            </div>
          </div>
        </div>

        {/* Search Filter Bar */}
        <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm mb-6">
          <div className="relative max-w-md">
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-3.5" />
            <input
              type="text"
              placeholder="Search by Name, Roll No, Email, or Section..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2 rounded-xl border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
          </div>
        </div>

        {/* Table of Students */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-100 text-[11px] font-extrabold uppercase text-slate-500 tracking-wider">
                  <th className="py-3.5 px-6">Roll / Register No</th>
                  <th className="py-3.5 px-6">Student Name</th>
                  <th className="py-3.5 px-6">College Email</th>
                  <th className="py-3.5 px-6">Department</th>
                  <th className="py-3.5 px-6">Section</th>
                  <th className="py-3.5 px-6">Phone Number</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-xs font-medium text-slate-700">
                {filteredStudents.length > 0 ? (
                  filteredStudents.map((s) => (
                    <tr key={s.id} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-3.5 px-6 font-mono font-bold text-[#082849]">{s.registerNo}</td>
                      <td className="py-3.5 px-6 font-bold text-slate-900">{s.name}</td>
                      <td className="py-3.5 px-6 text-slate-600">{s.email}</td>
                      <td className="py-3.5 px-6 font-semibold text-slate-800">{s.department}</td>
                      <td className="py-3.5 px-6 font-bold text-blue-700">Section {s.section}</td>
                      <td className="py-3.5 px-6 text-slate-500">{s.phone || 'N/A'}</td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={6} className="py-12 text-center text-slate-400">
                      <Database className="w-8 h-8 mx-auto mb-2 opacity-50" />
                      <p className="text-sm font-semibold">No student records found in database</p>
                      <p className="text-xs text-slate-400 mt-1">Upload an Excel sheet in Generator or connect Supabase DB to populate records.</p>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

      </div>
    </div>
  );
}
