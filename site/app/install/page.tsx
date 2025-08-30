import type { Metadata } from 'next'
import { redirect } from 'next/navigation'
import { createClient } from '@/lib/supabase/server'
import InstallContent from './InstallContent'

export const metadata: Metadata = {
  title: 'Install CellPilot - Google Sheets Add-on',
  description: 'Easy installation guide for CellPilot. Add powerful automation tools to your Google Sheets in minutes. Available on Google Workspace Marketplace.',
}

export default async function InstallPage() {
  const supabase = await createClient()
  
  const { data: { user } } = await supabase.auth.getUser()

  if (!user) {
    redirect('/auth?redirect=/install')
  }

  return <InstallContent />
}