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
    const { email, full_name, company, subscription_tier, is_admin } = body
    
    console.log('Creating user with data:', { email, full_name, company, subscription_tier, is_admin })
    
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
    
    // First check if user already exists with this email
    const { data: { users }, error: listError } = await serviceClient.auth.admin.listUsers()
    
    if (listError) {
      return NextResponse.json({ error: `Failed to check existing users: ${listError.message}` }, { status: 500 })
    }
    
    const existingUser = users.find(u => u.email === email)
    
    let userData
    
    if (existingUser) {
      console.log(`User already exists with email ${email}, using existing user`)
      userData = { user: existingUser }
    } else {
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
      
      userData = data
    }
    
    const data = userData
    
    // Create profile with additional settings
    if (data.user) {
      // Create profile with all fields that exist in the database
      const tier = subscription_tier || 'free'
      const profileData = {
        id: data.user.id,
        email: data.user.email || '',
        full_name: full_name || '',
        company: company || '',
        subscription_tier: tier,
        operations_used: 0,
        operations_limit: 
          tier === 'beta' ? 1000 :
          tier === 'pro' ? 5000 :
          tier === 'enterprise' ? 999999 : 25,
        email_verified: true,
        newsletter_subscribed: false,
        is_admin: is_admin || false,
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      }
      
      console.log('Creating profile with data:', profileData)
      
      // First check if profile already exists
      const { data: existingProfile } = await serviceClient
        .from('profiles')
        .select('id')
        .eq('id', data.user.id)
        .single()
      
      let newProfile
      
      if (existingProfile) {
        console.log('Profile already exists, updating instead')
        // Update existing profile
        const { data: updated, error: updateError } = await serviceClient
          .from('profiles')
          .update({
            email: profileData.email,
            full_name: profileData.full_name,
            company: profileData.company,
            subscription_tier: profileData.subscription_tier,
            operations_limit: profileData.operations_limit,
            email_verified: profileData.email_verified,
            is_admin: profileData.is_admin,
            updated_at: profileData.updated_at
          })
          .eq('id', data.user.id)
          .select()
          .single()
        
        if (updateError) {
          console.error('Error updating profile:', updateError)
          return NextResponse.json({ 
            error: `Failed to update existing profile: ${updateError.message}`, 
            details: updateError.message
          }, { status: 500 })
        }
        newProfile = updated
      } else {
        // Create new profile
        const { data: created, error: profileError } = await serviceClient
          .from('profiles')
          .insert(profileData)
          .select()
          .single()
        
        if (profileError) {
          console.error('Error creating profile:', profileError)
          console.error('Profile data that failed:', profileData)
          // Try to delete the auth user if profile creation fails
          await serviceClient.auth.admin.deleteUser(data.user.id)
          return NextResponse.json({ 
            error: `Failed to create user profile: ${profileError.message}`, 
            details: profileError.message,
            code: profileError.code,
            hint: profileError.hint
          }, { status: 500 })
        }
        newProfile = created
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