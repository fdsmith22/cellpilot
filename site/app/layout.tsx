import type { Metadata } from 'next'
import { Inter } from 'next/font/google'
import './globals.css'
import Header from '@/components/Header'
import Footer from '@/components/Footer'

const inter = Inter({ subsets: ['latin'] })

export const metadata: Metadata = {
  title: 'CellPilot - Spreadsheets Reimagined',
  description: 'Take control of your spreadsheet data simply and quickly with smart automation tools for Google Sheets.',
  keywords: 'Google Sheets, automation, spreadsheet, data cleaning, formula builder, add-on, remove duplicates, email alerts',
  authors: [{ name: 'CellPilot Ltd' }],
  metadataBase: new URL('https://cellpilot.app'),
  icons: {
    icon: [
      { url: '/favicon.svg', type: 'image/svg+xml' },
      { url: '/favicon.ico', sizes: 'any' },
    ],
    apple: '/apple-touch-icon.png',
    other: [
      {
        rel: 'android-chrome-192x192',
        url: '/android-chrome-192x192.png',
      },
      {
        rel: 'android-chrome-512x512',
        url: '/android-chrome-512x512.png',
      },
    ],
  },
  manifest: '/site.webmanifest',
  themeColor: '#2563eb',
  openGraph: {
    title: 'CellPilot - Spreadsheets Reimagined',
    description: 'Take control of your spreadsheet data simply and quickly with smart automation tools',
    url: 'https://cellpilot.app',
    siteName: 'CellPilot',
    type: 'website',
    images: [
      {
        url: '/logo/combined/horizontal-standard-200x60.svg',
        width: 200,
        height: 60,
        alt: 'CellPilot Logo',
      },
    ],
  },
  twitter: {
    card: 'summary_large_image',
    title: 'CellPilot - Spreadsheets Reimagined',
    description: 'Take control of your spreadsheet data simply and quickly with smart automation tools',
    images: ['/logo/combined/horizontal-standard-200x60.svg'],
  },
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en" className="scroll-smooth">
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover" />
      </head>
      <body className={inter.className}>
        <Header />
        {children}
        <Footer />
      </body>
    </html>
  )
}