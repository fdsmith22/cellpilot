import { createClient } from '@/lib/supabase/server'
import { NextResponse } from 'next/server'

export async function GET(request: Request) {
  const requestUrl = new URL(request.url)
  const code = requestUrl.searchParams.get('code')
  const token_hash = requestUrl.searchParams.get('token_hash')
  const type = requestUrl.searchParams.get('type')
  const origin = requestUrl.origin
  const next = requestUrl.searchParams.get('next') || '/dashboard'
  const error_description = requestUrl.searchParams.get('error_description')
  const error = requestUrl.searchParams.get('error')

  // Handle OAuth errors
  if (error) {
    console.error('OAuth error:', error, error_description)
    return NextResponse.redirect(`${origin}/auth?error=${encodeURIComponent(error_description || error)}`)
  }

  const supabase = await createClient()

  // Handle email confirmation with token_hash (magic links and email confirmations)
  if (token_hash && type) {
    try {
      const { data, error: verifyError } = await supabase.auth.verifyOtp({
        token_hash,
        type: type as any,
      })

      if (verifyError) {
        console.error('Error verifying OTP:', verifyError)
        return NextResponse.redirect(`${origin}/auth?error=${encodeURIComponent(verifyError.message)}`)
      }

      if (data.user) {
        // Update email_verified status
        await supabase
          .from('profiles')
          .update({ 
            email_verified: true,
            updated_at: new Date().toISOString()
          })
          .eq('id', data.user.id)

        // Check if new user
        const { data: profile } = await supabase
          .from('profiles')
          .select('created_at')
          .eq('id', data.user.id)
          .single()
        
        const isNewUser = profile?.created_at && 
          new Date(profile.created_at).getTime() > Date.now() - 3600000
        
        if (isNewUser || type === 'signup') {
          return NextResponse.redirect(`${origin}/auth?verified=true`)
        }
        
        return NextResponse.redirect(`${origin}${next}`)
      }
    } catch (err) {
      console.error('Token verification error:', err)
      return NextResponse.redirect(`${origin}/auth?error=${encodeURIComponent('Verification failed. Please try again.')}`)
    }
  }

  // Handle OAuth code exchange
  if (code) {
    try {
      const { data, error: exchangeError } = await supabase.auth.exchangeCodeForSession(code)
      
      if (exchangeError) {
        console.error('Error exchanging code for session:', exchangeError)
        // More specific error message
        const errorMessage = exchangeError.message || 'Unable to verify email'
        return NextResponse.redirect(`${origin}/auth?error=${encodeURIComponent(errorMessage)}`)
      }
      
      const { user } = data
      
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
          new Date(profile.created_at).getTime() > Date.now() - 3600000 // Created within last hour (increased from 1 minute)
        
        if (isNewUser) {
          // New user email confirmation - redirect to auth page with success message
          return NextResponse.redirect(`${origin}/auth?verified=true`)
        }
        
        // Existing user - redirect to dashboard
        return NextResponse.redirect(`${origin}${next}`)
      }
    } catch (err) {
      console.error('Callback error:', err)
      return NextResponse.redirect(`${origin}/auth?error=${encodeURIComponent('Verification failed. Please try again.')}`)
    }
  }

  // No code provided - redirect to home
  return NextResponse.redirect(`${origin}/`)
}