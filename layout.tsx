import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'StaphySearch — Virtuálna Nemocnica',
  description: 'Digitálne dvojča nemocnice · Analýza Staphylococcus aureus & epidermidis · Kramáre',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="sk">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
      </head>
      <body>{children}</body>
    </html>
  );
}
