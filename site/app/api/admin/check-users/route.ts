import { createClient } from '@/lib/supabase/server'
import { createClient as createServiceClient } from '@supabase/supabase-js'
import { NextResponse } from 'next/server'

export async function GET(request: Request) {
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
    
    // Get all profiles
    const { data: profiles, error: profilesError } = await serviceClient
      .from('profiles')
      .select('*')
      .order('created_at', { ascending: false })
    
    if (profilesError) {
      return NextResponse.json({ error: profilesError.message }, { status: 500 })
    }
    
    // Create a map for easier lookup
    const profileMap = new Map(profiles?.map(p => [p.id, p]) || [])
    
    // Analyze the data
    const analysis = {
      authUsersCount: authUsers.length,
      profilesCount: profiles?.length || 0,
      authUsers: authUsers.map(u => ({
        id: u.id,
        email: u.email,
        created_at: u.created_at,
        hasProfile: profileMap.has(u.id)
      })),
      profiles: profiles?.map(p => ({
        id: p.id,
        email: p.email,
        full_name: p.full_name,
        subscription_tier: p.subscription_tier,
        is_admin: p.is_admin,
        operations_limit: p.operations_limit
      })),
      missingProfiles: authUsers.filter(u => !profileMap.has(u.id)).map(u => ({
        id: u.id,
        email: u.email
      })),
      orphanProfiles: profiles?.filter(p => !authUsers.find(u => u.id === p.id)).map(p => ({
        id: p.id,
        email: p.email
      }))
    }
    
    return NextResponse.json(analysis)
    
  } catch (error) {
    console.error('Check users error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}