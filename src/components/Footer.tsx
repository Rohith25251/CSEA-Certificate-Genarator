import React from 'react';
import Link from 'next/link';

export const Footer: React.FC = () => {
  return (
    <footer className="bg-[#082849] text-white border-t-4 border-yellow-500">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-12">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          
          {/* Column 1: Association Branding */}
          <div className="space-y-4">
            <div className="flex items-center space-x-3">
              <div className="bg-white p-2 rounded-lg">
                <img src="/csea_logo.png" alt="CSEA Logo" className="h-8 w-auto" />
              </div>
              <span className="font-bold text-lg text-yellow-400">CSE Association</span>
            </div>
            <p className="text-sm text-blue-200 leading-relaxed">
              Computer Science Engineering Association (CSEA) Certificate Generation & Management Platform for workshops, symposiums, hackathons, and guest lectures.
            </p>
          </div>

          {/* Column 2: Quick Links */}
          <div>
            <h4 className="font-bold text-sm uppercase tracking-wider text-yellow-400 mb-4">Quick Portals</h4>
            <ul className="space-y-2 text-sm text-blue-100">
              <li><Link href="/dashboard" className="hover:text-yellow-400 transition-colors">Admin Dashboard</Link></li>
              <li><Link href="/dashboard/events" className="hover:text-yellow-400 transition-colors">Event Categories</Link></li>
              <li><Link href="/dashboard/events/evt-ml-workshop-2026/generator" className="hover:text-yellow-400 transition-colors">5-Step Certificate Generator</Link></li>
              <li><Link href="/profile" className="hover:text-yellow-400 transition-colors">Student Profile</Link></li>
              <li><Link href="/certificates" className="hover:text-yellow-400 transition-colors">My Certificates Gallery</Link></li>
            </ul>
          </div>

          {/* Column 3: Event Categories */}
          <div>
            <h4 className="font-bold text-sm uppercase tracking-wider text-yellow-400 mb-4">Supported Events</h4>
            <div className="flex flex-wrap gap-2">
              {['Workshop', 'Hackathon', 'Technical Symposium', 'Coding Contest', 'Paper Presentation', 'Guest Lecture', 'Webinar', 'Project Expo', 'Quiz'].map((cat) => (
                <span key={cat} className="text-xs bg-blue-900/60 border border-blue-700/50 text-blue-200 px-2.5 py-1 rounded-full">
                  {cat}
                </span>
              ))}
            </div>
          </div>

        </div>

        <div className="mt-8 pt-8 border-t border-blue-900 text-center text-xs text-blue-300">
          <p>© {new Date().getFullYear()} Computer Science Engineering Association (CSEA). All rights reserved.</p>
        </div>
      </div>
    </footer>
  );
};
