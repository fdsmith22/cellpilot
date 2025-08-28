import { NextRequest, NextResponse } from 'next/server'

// Check subscription status
export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url)
    const userId = searchParams.get('userId')
    
    if (!userId) {
      return NextResponse.json(
        { error: 'User ID required' },
        { status: 400 }
      )
    }

    // In production, fetch from database
    // For now, return mock data
    const subscription = {
      userId,
      plan: 'free',
      status: 'active',
      operationsUsed: 12,
      operationsLimit: 25,
      billingCycle: 'monthly',
      nextBillingDate: null,
      features: {
        duplicateRemoval: true,
        textStandardization: true,
        formulaBuilder: false,
        emailAutomation: false,
        calendarIntegration: false
      }
    }

    return NextResponse.json(subscription)
  } catch (error) {
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    )
  }
}

// Update subscription (upgrade/downgrade)
export async function PUT(request: NextRequest) {
  try {
    const body = await request.json()
    const { userId, plan, paymentMethod } = body

    // Verify authentication
    // In production, check auth token
    
    // Validate plan
    const validPlans = ['free', 'starter', 'professional', 'business']
    if (!validPlans.includes(plan)) {
      return NextResponse.json(
        { error: 'Invalid plan' },
        { status: 400 }
      )
    }

    // Process payment if upgrading
    if (plan !== 'free') {
      // Integrate with Stripe here
      // Processing payment for: { userId, plan, paymentMethod }
    }

    // Update subscription in database
    const updatedSubscription = {
      userId,
      plan,
      status: 'active',
      updatedAt: new Date().toISOString(),
      features: getFeaturesByPlan(plan)
    }

    return NextResponse.json({
      success: true,
      subscription: updatedSubscription
    })
  } catch (error) {
    console.error('Subscription update error:', error)
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    )
  }
}

function getFeaturesByPlan(plan: string) {
  const features = {
    free: {
      operationsLimit: 25,
      duplicateRemoval: true,
      textStandardization: true,
      formulaBuilder: false,
      emailAutomation: false,
      calendarIntegration: false,
      prioritySupport: false
    },
    starter: {
      operationsLimit: 500,
      duplicateRemoval: true,
      textStandardization: true,
      formulaBuilder: true,
      emailAutomation: false,
      calendarIntegration: false,
      prioritySupport: false
    },
    professional: {
      operationsLimit: -1, // unlimited
      duplicateRemoval: true,
      textStandardization: true,
      formulaBuilder: true,
      emailAutomation: true,
      calendarIntegration: true,
      prioritySupport: true
    },
    business: {
      operationsLimit: -1, // unlimited
      duplicateRemoval: true,
      textStandardization: true,
      formulaBuilder: true,
      emailAutomation: true,
      calendarIntegration: true,
      prioritySupport: true,
      teamFeatures: true,
      sso: true
    }
  }
  
  return features[plan as keyof typeof features] || features.free
}
