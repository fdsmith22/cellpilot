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

  // Get user profile to check beta access
  const { data: profile } = await supabase
    .from('profiles')
    .select('beta_access, beta_requested_at, beta_approved_at, subscription_tier')
    .eq('id', user.id)
    .single()

  return <InstallContent user={user} profile={profile} />
}