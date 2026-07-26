'use client';

import React, { useState } from 'react';
import { CheckCircle2, Edit2, Check } from 'lucide-react';

export default function ProfilePage() {
  const [isEditing, setIsEditing] = useState(false);
  const [profile, setProfile] = useState({
    name: 'Preethika sri K',
    registerNo: '25CSR220',
    department: 'Computer Science and Engineering',
    yearOfStudy: 'II Year (4th Sem)',
    section: 'D',
    email: 'preethikasrik.25cse@kongu.edu',
    phone: '9566522117',
  });

  // Custom Email Branding State
  const [emailLogoUrl, setEmailLogoUrl] = useState('/csea_logo.png');
  const [emailHeroImageUrl, setEmailHeroImageUrl] = useState('/hero.png');
  const [savedSettingsMsg, setSavedSettingsMsg] = useState('');

  return (
    <div className="min-h-screen bg-slate-50 py-10">
      <div className="max-w-5xl mx-auto px-4 sm:px-6 lg:px-8 space-y-8">
        
        {/* Profile Card Header */}
        <div className="bg-[#082849] text-white rounded-3xl p-8 shadow-xl relative overflow-hidden border-b-4 border-yellow-500">
          <div className="absolute right-0 top-0 opacity-15 pointer-events-none transform translate-x-10 -translate-y-10">
            <img src="/csea_logo.png" alt="CSEA" className="h-80 w-auto" />
          </div>

          <div className="relative z-10 flex flex-col md:flex-row items-center md:items-start space-y-6 md:space-y-0 md:space-x-8">
            <div className="w-24 h-24 bg-gradient-to-tr from-yellow-400 to-amber-500 text-[#082849] font-black text-3xl rounded-2xl flex items-center justify-center shadow-lg border-4 border-white/20">
              {profile.name.substring(0, 2).toUpperCase()}
            </div>

            <div className="text-center md:text-left flex-grow">
              <div className="flex flex-col md:flex-row md:items-center justify-between">
                <div>
                  <h1 className="text-2xl font-black">{profile.name}</h1>
                  <p className="text-blue-200 text-xs font-semibold mt-1">{profile.department} ({profile.yearOfStudy})</p>
                </div>
                
                <button
                  onClick={() => setIsEditing(!isEditing)}
                  className="mt-4 md:mt-0 flex items-center space-x-2 bg-yellow-400 hover:bg-yellow-300 text-[#082849] px-5 py-2.5 rounded-xl font-bold text-xs transition-colors shadow-md"
                >
                  <Edit2 className="w-4 h-4" />
                  <span>{isEditing ? 'Editing Profile' : 'Edit Profile'}</span>
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Association Logo Asset */}
        <div className="p-6 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
          <div className="flex items-center justify-between">
            <div>
              <h4 className="font-bold text-slate-900 text-sm">Association Logo Asset</h4>
              <p className="text-xs text-slate-500">Displayed in site header, email header, and certificate top bar</p>
            </div>
            <span className="font-mono text-xs text-blue-700 font-bold bg-blue-50 px-3 py-1 rounded-lg border border-blue-200">
              csea_logo.png
            </span>
          </div>
        </div>

        {/* Custom Email & Certificate Branding Card */}
        <div className="bg-white p-8 rounded-3xl border border-slate-200 shadow-xl space-y-6">
          <div>
            <h2 className="text-2xl font-extrabold text-[#082849]">Custom Email & Certificate Branding Settings</h2>
            <p className="text-xs text-slate-500 mt-1">
              Configure custom logo and hero banner assets for emails and certificate headers.
            </p>
          </div>

          {savedSettingsMsg && (
            <div className="p-3 bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-bold rounded-xl flex items-center space-x-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-600" />
              <span>{savedSettingsMsg}</span>
            </div>
          )}

          <form onSubmit={(e) => { e.preventDefault(); setSavedSettingsMsg('Branding settings saved successfully!'); setTimeout(() => setSavedSettingsMsg(''), 3000); }} className="space-y-6">
            
            {/* Logo Setting */}
            <div className="p-6 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-slate-900 text-sm">Association Logo Asset</h4>
                  <p className="text-xs text-slate-500">Displayed in site header, email header, and certificate top bar</p>
                </div>
                <span className="font-mono text-xs text-blue-700 font-bold bg-blue-50 px-3 py-1 rounded-lg border border-blue-200">
                  csea_logo.png
                </span>
              </div>

              <input
                type="text"
                value={emailLogoUrl}
                onChange={(e) => setEmailLogoUrl(e.target.value)}
                className="w-full p-3 bg-white border border-slate-300 rounded-xl text-xs font-mono text-slate-800"
              />
            </div>

            {/* Hero Banner Setting */}
            <div className="p-6 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-slate-900 text-sm">Hero Banner Asset</h4>
                  <p className="text-xs text-slate-500">Displayed in main hero section and email body</p>
                </div>
                <span className="font-mono text-xs text-blue-700 font-bold bg-blue-50 px-3 py-1 rounded-lg border border-blue-200">
                  hero.png
                </span>
              </div>

              <input
                type="text"
                value={emailHeroImageUrl}
                onChange={(e) => setEmailHeroImageUrl(e.target.value)}
                className="w-full p-3 bg-white border border-slate-300 rounded-xl text-xs font-mono text-slate-800"
              />
            </div>

            <div className="flex justify-end">
              <button
                type="submit"
                className="bg-[#082849] hover:bg-[#0c3f6e] text-yellow-400 font-extrabold px-6 py-3 rounded-xl text-xs uppercase tracking-wider shadow-md transition-colors"
              >
                Save Branding Configuration
              </button>
            </div>

          </form>
        </div>

      </div>
    </div>
  );
}
