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
      // First create the basic profile with only the fields from the trigger
      const basicProfileData = {
        id: data.user.id,
        email: data.user.email || '',
        full_name: full_name || '',
        company: company || '',
        newsletter_subscribed: false,
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      }
      
      // Insert basic profile first
      const { error: basicInsertError } = await serviceClient
        .from('profiles')
        .insert(basicProfileData)
      
      if (basicInsertError) {
        console.error('Error creating basic profile:', basicInsertError)
        // Try to delete the auth user if profile creation fails
        await serviceClient.auth.admin.deleteUser(data.user.id)
        return NextResponse.json({ 
          error: 'Failed to create user profile', 
          details: basicInsertError.message 
        }, { status: 500 })
      }
      
      // Now update with additional fields that might exist
      const updateData: any = {
        subscription_tier: subscription_tier || 'free',
        is_admin: is_admin || false,
        beta_access: beta_access || false,
        updated_at: new Date().toISOString()
      }
      
      // Add optional fields if they're supposed to be set
      if (beta_access) {
        updateData.beta_approved_at = new Date().toISOString()
      }
      
      // Add operations fields if they exist in the table
      updateData.operations_limit = 
        subscription_tier === 'beta' ? 1000 :
        subscription_tier === 'pro' ? 5000 :
        subscription_tier === 'enterprise' ? 999999 : 25
      updateData.operations_used = 0
      updateData.email_verified = true
      
      const { data: newProfile, error: updateError } = await serviceClient
        .from('profiles')
        .update(updateData)
        .eq('id', data.user.id)
        .select()
        .single()
      
      if (updateError) {
        console.error('Error updating profile with additional fields:', updateError)
        // Profile was created but couldn't be updated - return partial success
        return NextResponse.json({ 
          success: true, 
          user: basicProfileData,
          warning: 'Profile created but some fields could not be set'
        })
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