import { createClient } from '@/lib/supabase/server'
import { NextResponse } from 'next/server'

export async function POST(request: Request) {
  try {
    const supabase = await createClient()
    
    // Get the current user
    const { data: { user } } = await supabase.auth.getUser()
    
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }
    
    const now = new Date().toISOString()
    
    // Just submit the beta request (no auto-approval)
    const { data, error } = await supabase
      .from('profiles')
      .upsert({
        id: user.id,
        email: user.email,
        beta_requested_at: now,
        beta_notes: 'Requested via dashboard - pending admin approval',
        updated_at: now
      }, {
        onConflict: 'id'
      })
      .select()
      .single()
    
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