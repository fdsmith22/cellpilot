import { NextResponse } from 'next/server'
import { createClient } from '@/lib/supabase/server'

export async function POST(request: Request) {
  try {
    const { email } = await request.json()
    
    if (!email) {
      return NextResponse.json({ 
        hasBeta: false, 
        error: 'Email required' 
      }, { status: 400 })
    }

    const supabase = await createClient()
    
    // Use the SQL function to check beta access
    const { data, error } = await supabase
      .rpc('check_beta_access_by_email', { 
        user_email: email.toLowerCase() 
      })
    
    if (error) {
      console.error('Error checking beta access:', error)
      return NextResponse.json({ 
        hasBeta: false,
        error: 'Failed to check beta access' 
      })
    }
    
    if (!data || !data.found) {
      return NextResponse.json({ 
        hasBeta: false,
        error: 'User not found. Please sign up at cellpilot.io' 
      })
    }

    return NextResponse.json({ 
      hasBeta: data.hasBeta,
      subscriptionTier: data.subscriptionTier
    })

  } catch (error) {
    console.error('Error checking beta access:', error)
    return NextResponse.json({ 
      hasBeta: false, 
      error: 'Internal server error' 
    }, { status: 500 })
  }
}

// Allow CORS for Apps Script
export async function OPTIONS() {
  return new NextResponse(null, {
    status: 200,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    },
  })
}