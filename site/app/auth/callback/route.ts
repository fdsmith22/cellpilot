import { createClient } from '@/lib/supabase/server'
import { NextResponse } from 'next/server'

export async function GET(request: Request) {
  const requestUrl = new URL(request.url)
  const code = requestUrl.searchParams.get('code')
  const origin = requestUrl.origin
  const next = requestUrl.searchParams.get('next') || '/dashboard'

  if (code) {
    const supabase = await createClient()
    const { data: { user }, error } = await supabase.auth.exchangeCodeForSession(code)
    
    if (error) {
      // Redirect to auth page with error
      return NextResponse.redirect(`${origin}/auth?error=Unable to verify email`)
    }
    
    // Update email_verified status in profiles table
    if (user) {
      await supabase
        .from('profiles')
        .update({ 
          email_verified: true,
          updated_at: new Date().toISOString()
        })
        .eq('id', user.id)
      
      // Check if this is email confirmation (first time)
      const { data: profile } = await supabase
        .from('profiles')
        .select('created_at')
        .eq('id', user.id)
        .single()
      
      const isNewUser = profile?.created_at && 
        new Date(profile.created_at).getTime() > Date.now() - 60000 // Created within last minute
      
      if (isNewUser) {
        // New user email confirmation - redirect to dashboard with welcome message
        return NextResponse.redirect(`${origin}/dashboard?welcome=true`)
      }
    }
  }

  // URL to redirect to after sign in process completes
  return NextResponse.redirect(`${origin}${next}`)