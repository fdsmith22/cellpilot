import { createClient } from '@/lib/supabase/server'
import { createClient as createServiceClient } from '@supabase/supabase-js'
import { NextResponse } from 'next/server'

export async function DELETE(request: Request) {
  try {
    const supabase = await createClient()
    
    // Get the current user
    const { data: { user } } = await supabase.auth.getUser()
    
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    
    // Get confirmation from request body (optional)
    const body = await request.json().catch(() => ({}))
    const { confirmEmail } = body
    
    // Optional: Verify email matches for extra safety
    if (confirmEmail && confirmEmail !== user.email) {
      return NextResponse.json({ 
        error: 'Email confirmation does not match' 
      }, { status: 400 })
    }
    
    console.log(`User ${user.email} (${user.id}) is deleting their account`)
    
    // Create service client for deletion operations
    const serviceClient = createServiceClient(
      process.env.NEXT_PUBLIC_SUPABASE_URL!,
      process.env.SUPABASE_SERVICE_ROLE_KEY!,
      {
        auth: {
          autoRefreshToken: false,
          persistSession: false
        }
      }
    )
    
    // Delete profile first (in case auth deletion fails, we can recover)
    const { error: profileError } = await serviceClient
      .from('profiles')
      .delete()
      .eq('id', user.id)
    
    if (profileError) {
      console.error('Error deleting profile:', profileError)
      // Continue anyway - profile might not exist
    }
    
    // Delete any other user data (installations, email_preferences, etc.)
    // Add these if you have other tables with user data
    const { error: installationsError } = await serviceClient
      .from('installations')
      .delete()
      .eq('user_id', user.id)
    
    if (installationsError) {
      console.log('No installations to delete or error:', installationsError.message)
    }
    
    const { error: emailPrefsError } = await serviceClient
      .from('email_preferences')
      .delete()
      .eq('id', user.id)
    
    if (emailPrefsError) {
      console.log('No email preferences to delete or error:', emailPrefsError.message)
    }
    
    // Finally, delete the auth user
    const { error: authError } = await serviceClient.auth.admin.deleteUser(user.id)
    
    if (authError) {
      console.error('Error deleting auth user:', authError)
      return NextResponse.json({ 
        error: 'Failed to delete account completely. Please contact support.', 
        details: authError.message 
      }, { status: 500 })
    }
    
    // Sign out the user locally (they're already deleted from auth)
    await supabase.auth.signOut()
    
    return NextResponse.json({ 
      success: true, 
      message: 'Account successfully deleted. You can sign up again with the same email if you wish.' 
    })
    
  } catch (error) {
    console.error('Delete account error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}