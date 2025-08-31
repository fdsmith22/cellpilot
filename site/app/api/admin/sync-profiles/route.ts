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
    
    // Get all auth users
    const { data: { users: authUsers }, error: authError } = await serviceClient.auth.admin.listUsers()
    
    if (authError) {
      return NextResponse.json({ error: authError.message }, { status: 500 })
    }
    
    // Get all existing profiles
    const { data: existingProfiles } = await supabase
      .from('profiles')
      .select('id')
    
    const existingIds = new Set(existingProfiles?.map(p => p.id) || [])
    
    // Find users without profiles
    const missingProfiles = authUsers.filter(u => !existingIds.has(u.id))
    
    console.log(`Found ${missingProfiles.length} users without profiles`)
    
    // Create profiles for missing users
    if (missingProfiles.length > 0) {
      const profilesToCreate = missingProfiles.map(user => {
        console.log(`Creating profile for user: ${user.email}`)
        // Only include fields that actually exist in the profiles table
        return {
          id: user.id,
          email: user.email || '',
          full_name: user.user_metadata?.full_name || '',
          company: user.user_metadata?.company || '',
          subscription_tier: 'free',
          operations_used: 0,
          operations_limit: 25,
          email_verified: !!user.email_confirmed_at,
          newsletter_subscribed: user.user_metadata?.newsletter_subscribed || true,
          is_admin: false,
          created_at: user.created_at || new Date().toISOString(),
          updated_at: new Date().toISOString()
        }
      })
      
      console.log('Profiles to create:', JSON.stringify(profilesToCreate, null, 2))
      
      // Try to insert one by one to identify which user causes the error
      const results = []
      const errors = []
      
      for (const profile of profilesToCreate) {
        // First check if profile already exists (in case our check missed it)
        const { data: existingProfile } = await serviceClient
          .from('profiles')
          .select('*')
          .eq('id', profile.id)
          .single()
        
        if (existingProfile) {
          console.log(`Profile already exists for ${profile.email}, updating instead`)
          // Update existing profile
          const { data: updatedProfile, error: updateError } = await serviceClient
            .from('profiles')
            .update({
              email: profile.email,
              full_name: profile.full_name || existingProfile.full_name,
              company: profile.company || existingProfile.company,
              subscription_tier: existingProfile.subscription_tier || 'free',
              operations_limit: existingProfile.operations_limit || 25,
              updated_at: new Date().toISOString()
            })
            .eq('id', profile.id)
            .select()
            .single()
          
          if (updateError) {
            console.error(`Failed to update profile for ${profile.email}:`, updateError)
            errors.push({
              email: profile.email,
              error: `Update failed: ${updateError.message}`,
              code: updateError.code,
              hint: updateError.hint
            })
          } else {
            results.push(updatedProfile)
          }
        } else {
          // Try to insert new profile
          const { data, error: insertError } = await serviceClient
            .from('profiles')
            .insert(profile)
            .select()
            .single()
          
          if (insertError) {
            console.error(`Failed to create profile for ${profile.email}:`, insertError)
            errors.push({
              email: profile.email,
              error: insertError.message,
              code: insertError.code,
              hint: insertError.hint,
              details: insertError.details
            })
          } else {
            results.push(data)
          }
        }
      }
      
      if (errors.length > 0) {
        return NextResponse.json({ 
          error: 'Some profiles failed to create', 
          details: errors,
          succeeded: results.length,
          failed: errors.length,
          message: `Created ${results.length} profiles, ${errors.length} failed`
        }, { status: 207 }) // 207 Multi-Status
      }
      
      // All succeeded
      return NextResponse.json({ 
        success: true, 
        synced: results.length,
        message: `Successfully synced ${results.length} profiles`
      })
    }
    
    // No profiles to sync
    return NextResponse.json({ 
      success: true, 
      synced: 0,
      message: 'No missing profiles found - all users have profiles'
    })
    
  } catch (error) {
    console.error('Sync error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}