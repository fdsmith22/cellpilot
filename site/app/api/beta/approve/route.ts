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
    
    // Check if profile exists first
    const { data: existingProfile } = await serviceClient
      .from('profiles')
      .select('id')
      .eq('id', user.id)
      .single()
    
    const now = new Date().toISOString()
    let data, error
    
    if (!existingProfile) {
      // Create profile with beta access if it doesn't exist
      const result = await serviceClient
        .from('profiles')
        .insert({
          id: user.id,
          email: user.email,
          beta_requested_at: now,
          beta_approved_at: now,
          beta_access: true,
          beta_notes: 'Auto-approved via API',
          subscription_tier: 'beta',
          created_at: now,
          updated_at: now
        })
        .select()
        .single()
      
      data = result.data
      error = result.error
    } else {
      // Update existing profile
      const result = await serviceClient
        .from('profiles')
        .update({ 
          beta_requested_at: now,
          beta_approved_at: now,
          beta_access: true,
          beta_notes: 'Auto-approved via API',
          subscription_tier: 'beta',
          updated_at: now
        })
        .eq('id', user.id)
        .select()
        .single()
      
      data = result.data
      error = result.error
    }
    
    if (error) {
      console.error('Error approving beta access:', error)
      return NextResponse.json({ 
        error: 'Failed to approve beta access', 
        details: error.message 
      }, { status: 500 })
    }
    
    // Return the script URL for installation
    const scriptUrl = 'https://script.google.com/d/1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O/edit'
    
    return NextResponse.json({ 
      success: true, 
      message: 'Beta access approved!',
      scriptUrl,
      installInstructions: [
        'Click the link to open the Apps Script editor',
        'Click Deploy → Test deployments',
        'Click Install button',
        'Click Done',
        'Open any Google Sheet and find CellPilot in Extensions menu'
      ]
    })
    
  } catch (error) {
    console.error('Beta approval error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}