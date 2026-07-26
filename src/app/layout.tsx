import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'CSE Association Certificate System',
  description: 'Dynamic Certificate Generation, Emailing, and Management System for Computer Science Engineering Association',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body className="min-h-screen bg-slate-50 font-sans">
        <main>{children}</main>
      </body>
    </html>
  );
}
