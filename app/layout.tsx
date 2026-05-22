import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'BOP CNX Title Invoice Generator',
  description: 'Generate CNX-format invoices from BOP Abstract work logs',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
