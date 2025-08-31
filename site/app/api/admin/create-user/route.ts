import { createClient } from '@/lib/supabase/server'
import { createClient as createServiceClient } from '@supabase/supabase-js'
import { NextResponse } from 'next/server'

export async function POST(request: Request) {
  try {
    const supabase = await createClient()
    
    // Check if user is admin
    const { data: { user } } = await supabase.auth.getUser()
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    
    const { data: profile } = await supabase
      .from('profiles')
      .select('is_admin')
      .eq('id', user.id)
      .single()
    
    if (!profile?.is_admin) {
      return NextResponse.json({ error: 'Unauthorized - Admin only' }, { status: 403 })
    }
    
    // Get request body
    const body = await request.json()
    const { email, full_name, company, subscription_tier, is_admin, beta_access } = body
    
    // Create service client for admin operations
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
    
    // Create user via admin API
    const { data, error } = await serviceClient.auth.admin.createUser({
      email,
      email_confirm: true,
      user_metadata: {
        full_name,
        company
      }
    })
    
    if (error) {
      console.error('Error creating user:', error)
      return NextResponse.json({ error: error.message }, { status: 400 })
    }
    
    // Create profile with additional settings
    if (data.user) {
      // Create profile with all fields that exist in the database
      const profileData = {
        id: data.user.id,
        email: data.user.email || '',
        full_name: full_name || '',
        company: company || '',
        subscription_tier: subscription_tier || 'free',
        operations_used: 0,
        operations_limit: 
          subscription_tier === 'beta' ? 1000 :
          subscription_tier === 'pro' ? 5000 :
          subscription_tier === 'enterprise' ? 999999 : 25,
        email_verified: true,
        newsletter_subscribed: false,
        is_admin: is_admin || false,
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      }
      
      const { data: newProfile, error: profileError } = await serviceClient
        .from('profiles')
        .insert(profileData)
        .select()
        .single()
      
      if (profileError) {
        console.error('Error creating profile:', profileError)
        // Try to delete the auth user if profile creation fails
        await serviceClient.auth.admin.deleteUser(data.user.id)
        return NextResponse.json({ 
          error: 'Failed to create user profile', 
          details: profileError.message 
        }, { status: 500 })
      }
      
      return NextResponse.json({ 
        success: true, 
        user: newProfile 
      })
    }
    
    return NextResponse.json({ error: 'User creation failed' }, { status: 500 })
    
  } catch (error) {
    console.error('Create user error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}