'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { createClient } from '@/lib/supabase/client'

interface BetaAccessCardProps {
  profile: any
  userId: string
}

export default function BetaAccessCard({ profile, userId }: BetaAccessCardProps) {
  const [loading, setLoading] = useState(false)
  const [hasBetaAccess, setHasBetaAccess] = useState(profile?.subscription_tier === 'beta')
  const router = useRouter()
  const supabase = createClient()

  const activateBetaAccess = async () => {
    setLoading(true)
    try {
      // Simple update to set user as beta tier
      const { error } = await supabase
        .from('profiles')
        .update({ 
          subscription_tier: 'beta',
          updated_at: new Date().toISOString()
        })
        .eq('id', userId)

      if (error) throw error

      setHasBetaAccess(true)
      router.refresh()
    } catch (error: any) {
      console.error('Error activating beta access:', error)
      alert('Failed to activate beta access. Please try again.')
    } finally {
      setLoading(false)
    }
  }

  const installCellPilot = () => {
    // Open the Beta Installer Web App (v1.6 - uses library v11)
    window.open(
      'https://script.google.com/macros/s/AKfycbwZn2ThA3Lb5Jiqdnv3AZS9yv5vlXauB6_0ogzO1_BRLpBXE7pFolhvl0Bx3f6cIneE/exec',
      '_blank'
    )
  }

  if (hasBetaAccess) {
    return (
      <div className="glass-card rounded-2xl p-8 bg-gradient-to-br from-green-50 to-emerald-50 border-2 border-green-200">
        <div className="flex items-start justify-between">
          <div>
            <h2 className="text-xl font-semibold text-green-900 mb-2">Beta Access Active!</h2>
            <p className="text-green-700 mb-4">
              You have full access to all CellPilot features. Click below to install.
            </p>
            <p className="text-sm text-green-600 mb-6">
              Subscription: Beta Tier (Full Access)
            </p>
          </div>
          <svg className="w-8 h-8 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
        </div>
        
        <div className="space-y-4">
          <button
            onClick={installCellPilot}
            className="w-full px-6 py-3 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors font-medium"
          >
            Install CellPilot Now
            <svg className="inline-block w-4 h-4 ml-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" />
            </svg>
          </button>
          
          <details className="group">
            <summary className="cursor-pointer text-sm text-green-700 hover:text-green-800 font-medium">
              📋 Installation Instructions (click to expand)
            </summary>
            <div className="mt-3 space-y-3 text-sm text-green-600 bg-white/50 rounded-lg p-4">
              <p className="font-medium text-green-800">Quick Install Steps:</p>
              <ol className="space-y-2 ml-4">
                <li>1. Click "Install CellPilot Now" above</li>
                <li>2. Sign in with your Google account (same email as CellPilot signup)</li>
                <li>3. Copy the installation code shown</li>
                <li>4. Open any Google Sheet</li>
                <li>5. Go to <strong>Extensions</strong> → <strong>Apps Script</strong></li>
                <li>6. Delete existing code and paste the installation code</li>
                <li>7. Add the CellPilot library (instructions provided)</li>
                <li>8. Save and refresh your Google Sheet</li>
                <li>9. Find <strong>CellPilot</strong> in the <strong>Extensions</strong> menu!</li>
              </ol>
              <p className="text-xs text-amber-600 mt-3">
                💡 Important: Use the same Google account email that you used to sign up for CellPilot
              </p>
            </div>
          </details>
        </div>
      </div>
    )
  }


  return (
    <div className="glass-card rounded-2xl p-8 bg-gradient-to-br from-primary-50 to-blue-50 border-2 border-primary-200">
      <div className="flex items-start justify-between">
        <div>
          <h2 className="text-xl font-semibold text-primary-900 mb-2">Start Beta Testing</h2>
          <p className="text-primary-700 mb-4">
            Get instant access to all CellPilot features during our beta phase!
          </p>
          <ul className="space-y-2 text-sm text-primary-600 mb-6">
            <li className="flex items-center">
              <svg className="w-4 h-4 mr-2 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
              Full access to all features
            </li>
            <li className="flex items-center">
              <svg className="w-4 h-4 mr-2 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
              No payment required during beta
            </li>
            <li className="flex items-center">
              <svg className="w-4 h-4 mr-2 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
              Help shape the product with your feedback
            </li>
          </ul>
        </div>
        <svg className="w-8 h-8 text-primary-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
        </svg>
      </div>
      
      <button
        onClick={activateBetaAccess}
        disabled={loading}
        className="w-full px-6 py-3 bg-primary-600 text-white rounded-lg hover:bg-primary-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed"
      >
        {loading ? 'Activating...' : 'Activate Beta Access'}
      </button>
    </div>
  )
}