import type { Metadata } from 'next'
import { Inter } from 'next/font/google'
import './globals.css'

const inter = Inter({ subsets: ['latin'] })

export const metadata: Metadata = {
  title: 'CellPilot - Google Sheets Automation Made Simple',
  description: 'Transform your Google Sheets workflow with powerful automation tools, data cleaning, and formula generation.',
  keywords: 'Google Sheets, automation, spreadsheet, data cleaning, formula builder, add-on',
  authors: [{ name: 'CellPilot Team' }],
  openGraph: {
    title: 'CellPilot - Google Sheets Automation',
    description: 'Transform your Google Sheets workflow with powerful automation tools',
    url: 'https://cellpilot.io',
    siteName: 'CellPilot',
    type: 'website',
  },
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <body className={inter.className}>{children}</body>
    </html>
  )
}