'use client';

import React, { useState, useEffect } from 'react';
import Link from 'next/link';
import { StoreManager } from '@/lib/store';
import { CseaEvent } from '@/lib/db';
import { Sparkles, Filter, Calendar, ArrowRight } from 'lucide-react';

const CATEGORIES = [
  { key: 'all', label: 'All Categories' },
  { key: 'workshop', label: 'Workshop' },
  { key: 'hackathon', label: 'Hackathon' },
  { key: 'technical_symposium', label: 'Technical Symposium' },
  { key: 'coding_contest', label: 'Coding Contest' },
  { key: 'paper_presentation', label: 'Paper Presentation' },
  { key: 'guest_lecture', label: 'Guest Lecture' },
  { key: 'webinar', label: 'Webinar' },
  { key: 'project_expo', label: 'Project Expo' },
  { key: 'quiz', label: 'Quiz' },
  { key: 'ideathon', label: 'Ideathon' },
  { key: 'other', label: 'Other' },
];

export default function EventsPage() {
  const [events, setEvents] = useState<CseaEvent[]>([]);
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');

  useEffect(() => {
    setEvents(StoreManager.getEvents());
  }, []);

  const filteredEvents = events.filter(event => {
    const matchesCategory = selectedCategory === 'all' || event.eventCategory === selectedCategory;
    const matchesSearch = event.eventName.toLowerCase().includes(searchQuery.toLowerCase()) ||
                          (event.description || '').toLowerCase().includes(searchQuery.toLowerCase());
    return matchesCategory && matchesSearch;
  });

  return (
    <div className="min-h-screen bg-slate-50 py-10">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        
        {/* Header */}
        <div className="flex flex-col md:flex-row md:items-center md:justify-between mb-8 pb-6 border-b border-slate-200">
          <div>
            <h1 className="text-3xl font-extrabold text-[#082849]">CSEA Events Directory</h1>
            <p className="text-sm text-slate-500 mt-1">
              Manage certificate templates and batch generation per event category
            </p>
          </div>

          <div className="mt-4 md:mt-0">
            <Link
              href="/"
              className="inline-flex items-center space-x-2 bg-gradient-to-r from-yellow-500 to-amber-500 text-[#082849] font-extrabold px-5 py-2.5 rounded-xl shadow-md hover:from-yellow-400 hover:to-amber-400 text-xs uppercase tracking-wider transition-all"
            >
              <Sparkles className="w-4 h-4" />
              <span>Launch Generator</span>
            </Link>
          </div>
        </div>

        {/* Category Filters Bar */}
        <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm mb-8">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            
            <div className="relative flex-grow max-w-md">
              <input
                type="text"
                placeholder="Search events by name..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-4 pr-10 py-2 rounded-xl border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
            </div>

            <div className="flex items-center space-x-2 overflow-x-auto pb-2 md:pb-0">
              <Filter className="w-4 h-4 text-slate-400 flex-shrink-0" />
              <div className="flex space-x-1.5 flex-nowrap">
                {CATEGORIES.slice(0, 6).map((cat) => (
                  <button
                    key={cat.key}
                    onClick={() => setSelectedCategory(cat.key)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition-colors ${
                      selectedCategory === cat.key
                        ? 'bg-[#082849] text-yellow-400 font-bold shadow-sm'
                        : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                    }`}
                  >
                    {cat.label}
                  </button>
                ))}
              </div>
            </div>

          </div>
        </div>

        {/* Events Cards Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {filteredEvents.map((event) => (
            <div
              key={event.id || event.eventId}
              className="bg-white rounded-2xl border border-slate-200 shadow-sm hover:shadow-md transition-shadow overflow-hidden flex flex-col justify-between"
            >
              <div className="p-6">
                <div className="flex items-center justify-between mb-3">
                  <span className="bg-blue-50 text-blue-700 text-xs font-extrabold uppercase px-2.5 py-1 rounded border border-blue-200">
                    {event.eventCategory.replace('_', ' ')}
                  </span>
                  <span className="text-xs text-slate-400 font-medium flex items-center">
                    <Calendar className="w-3.5 h-3.5 mr-1" />
                    {event.eventDate}
                  </span>
                </div>

                <h3 className="text-xl font-extrabold text-slate-900 mb-2">{event.eventName}</h3>
                <p className="text-sm text-slate-600 line-clamp-2 leading-relaxed mb-4">
                  {event.description}
                </p>

                <div className="bg-slate-50 p-3 rounded-xl border border-slate-100 space-y-1 text-xs">
                  <div className="flex justify-between">
                    <span className="text-slate-500 font-medium">Dataset Status:</span>
                    <span className="font-bold text-emerald-600">Excel Active (76 rows)</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-slate-500 font-medium">Template Tokens:</span>
                    <span className="font-mono text-blue-600 font-bold">&lt;&lt;Name&gt;&gt;, &lt;&lt;Roll Number&gt;&gt;</span>
                  </div>
                </div>
              </div>

              <div className="p-4 bg-slate-50/80 border-t border-slate-100 flex justify-between items-center">
                <span className="text-xs font-bold text-slate-600">
                  {event.totalParticipants || 76} Certificates
                </span>

                <Link
                  href="/"
                  className="inline-flex items-center space-x-1.5 bg-[#082849] hover:bg-[#0c3f6e] text-yellow-400 font-bold text-xs px-3.5 py-2 rounded-xl transition-colors"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  <span>Run Generator</span>
                </Link>
              </div>
            </div>
          ))}
        </div>

      </div>
    </div>
  );
}
