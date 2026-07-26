'use client';

import React from 'react';
import Link from 'next/link';
import { usePathname } from 'next/navigation';
import { Sparkles, FolderKanban, Users, User } from 'lucide-react';

export const Header: React.FC = () => {
  const pathname = usePathname();

  const navItems = [
    { href: '/', label: 'Generator', icon: Sparkles },
    { href: '/events', label: 'Events', icon: FolderKanban },
    { href: '/students', label: 'Students', icon: Users },
    { href: '/profile', label: 'Profile', icon: User },
  ];

  return (
    <header className="sticky top-0 z-50 bg-[#082849] backdrop-blur-md border-b-2 border-yellow-500 text-white shadow-lg">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-24">
          
          {/* Logo Branding */}
          <Link href="/" className="flex items-center space-x-3.5 group">
            <div className="w-14 h-14 bg-white p-1 rounded-2xl shadow-md border-2 border-indigo-500 group-hover:scale-105 transition-transform duration-200 flex items-center justify-center shrink-0">
              <img
                src="/csea_logo.png"
                alt="CSEA Logo"
                className="w-full h-full object-contain"
              />
            </div>
            <div>
              <div className="font-extrabold text-lg leading-tight tracking-wide text-white group-hover:text-yellow-400 transition-colors">
                CSE ASSOCIATION
              </div>
              <div className="text-xs text-yellow-400 font-mono font-bold tracking-wider">
                Certificate Platform
              </div>
            </div>
          </Link>

          {/* Clean 4 Navigation Links */}
          <nav className="flex items-center space-x-2">
            {navItems.map((item) => {
              const Icon = item.icon;
              const isActive = pathname === item.href;
              
              return (
                <Link
                  key={item.href}
                  href={item.href}
                  className={`flex items-center space-x-2 px-4 py-2 rounded-xl text-xs font-bold transition-all duration-200 ${
                    isActive
                      ? 'bg-yellow-500 text-[#082849] shadow-md font-extrabold'
                      : 'text-blue-100 hover:bg-white/10 hover:text-white'
                  }`}
                >
                  <Icon className="w-4 h-4" />
                  <span>{item.label}</span>
                </Link>
              );
            })}
          </nav>

        </div>
      </div>
    </header>
  );
};
