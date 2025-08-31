import { createClient } from '@/lib/supabase/server'
import { createClient as createServiceClient } from '@supabase/supabase-js'
import { NextResponse } from 'next/server'

export async function POST(request: Request) {
  try {
    const supabase = await createClient()
    
    // Get the current user and verify they're an admin
    const { data: { user } } = await supabase.auth.getUser()
    
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    
    // Check if user is admin
    const { data: profile } = await supabase
      .from('profiles')
      .select('email')
      .eq('id', user.id)
      .single()
    
    // Replace with your admin email
    if (profile?.email !== 'freddywsmith@gmail.com') {
      return NextResponse.json({ error: 'Admin access required' }, { status: 403 })
    }
    
    const { userId } = await request.json()
    
    if (!userId) {
      return NextResponse.json({ error: 'User ID required' }, { status: 400 })
    }
    
    // Create service client to bypass RLS
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
    
    const now = new Date().toISOString()
    
    // Approve beta access
    const { data, error } = await serviceClient
      .from('profiles')
      .update({
        beta_approved_at: now,
        beta_access: true,
        subscription_tier: 'beta',
        updated_at: now
      })
      .eq('id', userId)
      .select()
      .single()
    
    if (error) {
      console.error('Error approving beta access:', error)
      return NextResponse.json({ 
        error: 'Failed to approve beta access', 
        details: error.message 
      }, { status: 500 })
    }
    
    return NextResponse.json({ 
      success: true, 
      message: 'Beta access approved successfully',
      user: data
    })
    
  } catch (error) {
    console.error('Admin beta approval error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}