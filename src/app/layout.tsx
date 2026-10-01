import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'CSE Association Certificate System',
  description: 'Dynamic Certificate Generation, Emailing, and Management System for Computer Science Engineering Association',
  icons: {
    icon: '/csea_logo.png',
    shortcut: '/csea_logo.png',
    apple: '/csea_logo.png',
  },
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <head>
        <link rel="icon" href="/csea_logo.png" type="image/png" />
        <link rel="apple-touch-icon" href="/csea_logo.png" />
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no" />
      </head>
      <body className="min-h-screen bg-slate-50 font-sans antialiased overflow-x-hidden">
        <main className="w-full overflow-x-hidden">{children}</main>
      </body>
    </html>
  );
}
