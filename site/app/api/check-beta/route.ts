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
    
    // Check if user exists and has beta tier
    const { data: profile, error } = await supabase
      .from('profiles')
      .select('subscription_tier')
      .eq('email', email)
      .single()

    if (error || !profile) {
      return NextResponse.json({ 
        hasBeta: false,
        error: 'User not found. Please sign up at cellpilot.io' 
      })
    }

    return NextResponse.json({ 
      hasBeta: profile.subscription_tier === 'beta',
      subscriptionTier: profile.subscription_tier
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