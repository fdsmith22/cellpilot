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
    
    // Get email from request body
    const body = await request.json()
    const { email } = body
    
    if (!email) {
      return NextResponse.json({ error: 'Email required' }, { status: 400 })
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
    
    // Find user by email
    const { data: { users }, error: searchError } = await serviceClient.auth.admin.listUsers()
    
    if (searchError) {
      return NextResponse.json({ error: searchError.message }, { status: 500 })
    }
    
    const targetUser = users.find(u => u.email === email)
    
    if (!targetUser) {
      return NextResponse.json({ 
        message: 'No auth user found with that email',
        suggestion: 'User may have been partially deleted. Try signing up again.'
      })
    }
    
    // Delete from profiles first
    const { error: profileError } = await serviceClient
      .from('profiles')
      .delete()
      .eq('id', targetUser.id)
    
    if (profileError) {
      console.log('Profile deletion error (may not exist):', profileError)
    }
    
    // Delete from auth.users
    const { error: authError } = await serviceClient.auth.admin.deleteUser(targetUser.id)
    
    if (authError) {
      return NextResponse.json({ 
        error: 'Failed to delete auth user', 
        details: authError.message 
      }, { status: 500 })
    }
    
    return NextResponse.json({ 
      success: true, 
      message: `User ${email} completely removed from system. They can now sign up again.`
    })
    
  } catch (error) {
    console.error('Cleanup user error:', error)
    return NextResponse.json({ 
      error: 'Internal server error', 
      details: (error as Error).message 
    }, { status: 500 })
  }
}