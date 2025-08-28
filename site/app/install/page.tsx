import type { Metadata } from 'next'
import InstallContent from './InstallContent'

export const metadata: Metadata = {
  title: 'Install CellPilot - Google Sheets Add-on',
  description: 'Easy installation guide for CellPilot. Add powerful automation tools to your Google Sheets in minutes. Available on Google Workspace Marketplace.',
}

export default function InstallPage() {
  return <InstallContent />
}