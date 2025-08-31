'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'

interface BetaAccessCardProps {
  profile: any
  userId: string
}

export default function BetaAccessCard({ profile, userId }: BetaAccessCardProps) {
  const [loading, setLoading] = useState(false)
  const [requested, setRequested] = useState(!!profile?.beta_requested_at)
  const [approved, setApproved] = useState(!!profile?.beta_access)
  const router = useRouter()

  const requestBetaAccess = async () => {
    setLoading(true)
    try {
      // Use API endpoint to handle beta approval (bypasses RLS)
      const response = await fetch('/api/beta/approve', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
      })

      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.error || 'Failed to approve beta access')
      }

      console.log('Beta access granted:', data)
      setApproved(true)
      setRequested(true)
      
      // Refresh the page to update the UI
      router.refresh()
    } catch (error: any) {
      console.error('Error requesting beta access:', error)
      alert(`Failed to request beta access: ${error.message || 'Unknown error'}`)
    } finally {
      setLoading(false)
    }
  }

  const installCellPilot = () => {
    // Open the Apps Script project for test deployment installation
    window.open(
      'https://script.google.com/d/1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O/edit?showDeploy=true',
      '_blank'
    )
  }

  if (approved) {
    return (
      <div className="glass-card rounded-2xl p-8 bg-gradient-to-br from-green-50 to-emerald-50 border-2 border-green-200">
        <div className="flex items-start justify-between">
          <div>
            <h2 className="text-xl font-semibold text-green-900 mb-2">Beta Access Approved!</h2>
            <p className="text-green-700 mb-4">
              You have been approved for CellPilot beta access. Click below to install.
            </p>
            <p className="text-sm text-green-600 mb-6">
              Approved on: {profile.beta_approved_at ? new Date(profile.beta_approved_at).toLocaleDateString() : 'Recently'}
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
                <li>2. In the Apps Script editor that opens:
                  <ul className="ml-4 mt-1 text-xs">
                    <li>• Click <strong>Deploy</strong> → <strong>Test deployments</strong></li>
                    <li>• Click <strong>Install</strong> button</li>
                    <li>• Click <strong>Done</strong></li>
                  </ul>
                </li>
                <li>3. Open any Google Sheet</li>
                <li>4. Look for <strong>CellPilot</strong> in the <strong>Extensions</strong> menu</li>
                <li>5. Click any CellPilot menu item to authorize (first time only)</li>
              </ol>
              <p className="text-xs text-amber-600 mt-3">
                💡 Tip: You may need to refresh your Google Sheet for the menu to appear
              </p>
            </div>
          </details>
        </div>
      </div>
    )
  }

  if (requested) {
    return (
      <div className="glass-card rounded-2xl p-8 bg-gradient-to-br from-amber-50 to-yellow-50 border-2 border-amber-200">
        <div className="flex items-start justify-between">
          <div>
            <h2 className="text-xl font-semibold text-amber-900 mb-2">Beta Access Requested</h2>
            <p className="text-amber-700 mb-4">
              Your beta access request is being reviewed. We'll notify you once approved!
            </p>
            <p className="text-sm text-amber-600">
              Requested on: {profile.beta_requested_at ? new Date(profile.beta_requested_at).toLocaleDateString() : 'Just now'}
            </p>
          </div>
          <svg className="w-8 h-8 text-amber-500 animate-pulse" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
        </div>
        <p className="text-xs text-amber-600 mt-4">
          Typical approval time: Within 24 hours
        </p>
      </div>
    )
  }

  return (
    <div className="glass-card rounded-2xl p-8 bg-gradient-to-br from-primary-50 to-blue-50 border-2 border-primary-200">
      <div className="flex items-start justify-between">
        <div>
          <h2 className="text-xl font-semibold text-primary-900 mb-2">Get Beta Access</h2>
          <p className="text-primary-700 mb-4">
            Join the CellPilot beta program and get early access to all features!
          </p>
          <ul className="space-y-2 text-sm text-primary-600 mb-6">
            <li className="flex items-center">
              <svg className="w-4 h-4 mr-2 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
              All premium features unlocked
            </li>
            <li className="flex items-center">
              <svg className="w-4 h-4 mr-2 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
              Direct feedback channel to shape the product
            </li>
            <li className="flex items-center">
              <svg className="w-4 h-4 mr-2 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
              Special pricing when we launch
            </li>
          </ul>
        </div>
        <svg className="w-8 h-8 text-primary-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
        </svg>
      </div>
      
      <button
        onClick={requestBetaAccess}
        disabled={loading}
        className="w-full px-6 py-3 bg-primary-600 text-white rounded-lg hover:bg-primary-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed"
      >
        {loading ? 'Requesting...' : 'Request Beta Access'}
      </button>
    </div>
  )
}