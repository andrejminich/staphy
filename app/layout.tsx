import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Rezistencia Dashboard · Kramáre',
  description: 'Analýza MDR kmeňov S. aureus a S. epidermidis',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="sk">
      <body>{children}</body>
    </html>
  );
}
