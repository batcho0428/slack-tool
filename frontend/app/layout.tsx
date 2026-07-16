import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Slack Tool (Next.js)',
  description: 'Google Apps Script backend + Next.js frontend'
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja">
      <head>
        <link
          href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css"
          rel="stylesheet"
        />
      </head>
      <body>
        {children}
      </body>
    </html>
  );
}
