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
        // Only include fields that exist in the database
        return {
          id: user.id,
          email: user.email || '',
          full_name: user.user_metadata?.full_name || '',
          first_name: user.user_metadata?.first_name || '',
          surname: user.user_metadata?.surname || '',
          company: user.user_metadata?.company || '',
          newsletter_subscribed: user.user_metadata?.newsletter_subscribed || false,
          created_at: user.created_at || new Date().toISOString(),
          updated_at: new Date().toISOString()
        }
      })
      
      console.log('Profiles to create:', JSON.stringify(profilesToCreate, null, 2))
      
      // Try to insert one by one to identify which user causes the error
      const results = []
      const errors = []
      
      for (const profile of profilesToCreate) {
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
            code: insertError.code
          })
        } else {
          results.push(data)
        }
      }
      
      if (errors.length > 0) {
        return NextResponse.json({ 
          error: 'Some profiles failed to create', 
          details: errors,
          succeeded: results.length,
          failed: errors.length
        }, { status: 207 }) // 207 Multi-Status
      }
    }
    
    return NextResponse.json({ 
      success: true, 
      synced: missingProfiles.length,
      message: `Synced ${missingProfiles.length} missing profiles`
    })
    
  } catch (error) {
    console.error('Sync error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}