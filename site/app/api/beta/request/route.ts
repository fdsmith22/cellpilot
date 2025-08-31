import { createClient } from '@/lib/supabase/server'
import { createClient as createServiceClient } from '@supabase/supabase-js'
import { NextResponse } from 'next/server'

export async function POST(request: Request) {
  try {
    const supabase = await createClient()
    
    // Get the current user
    const { data: { user } } = await supabase.auth.getUser()
    
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    
    // Create service client to bypass RLS for profile updates
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
    
    // Check if profile exists
    const { data: existingProfile } = await serviceClient
      .from('profiles')
      .select('id')
      .eq('id', user.id)
      .single()
    
    let data, error
    
    if (!existingProfile) {
      // Create new profile with beta request
      const result = await serviceClient
        .from('profiles')
        .insert({
          id: user.id,
          email: user.email,
          beta_requested_at: now,
          beta_notes: 'Requested via dashboard - pending admin approval',
          created_at: now,
          updated_at: now
        })
        .select()
        .single()
      
      data = result.data
      error = result.error
    } else {
      // Update existing profile with beta request
      const result = await serviceClient
        .from('profiles')
        .update({
          beta_requested_at: now,
          beta_notes: 'Requested via dashboard - pending admin approval',
          updated_at: now
        })
        .eq('id', user.id)
        .select()
        .single()
      
      data = result.data
      error = result.error
    }
    
    if (error) {
      console.error('Error submitting beta request:', error)
      return NextResponse.json({ 
        error: 'Failed to submit beta request', 
        details: error.message 
      }, { status: 500 })
    }
    
    return NextResponse.json({ 
      success: true, 
      message: 'Beta access requested! You will be notified once approved by an admin.',
      requested: true
    })
    
  } catch (error) {
    console.error('Beta approval error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}