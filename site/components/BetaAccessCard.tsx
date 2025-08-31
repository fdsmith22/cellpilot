'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import { useRouter } from 'next/navigation'

interface BetaAccessCardProps {
  profile: any
  userId: string
}

export default function BetaAccessCard({ profile, userId }: BetaAccessCardProps) {
  const [loading, setLoading] = useState(false)
  const [requested, setRequested] = useState(profile?.beta_requested_at !== null)
  const [approved, setApproved] = useState(profile?.beta_access === true)
  const router = useRouter()
  const supabase = createClient()

  const requestBetaAccess = async () => {
    setLoading(true)
    try {
      const { error } = await supabase
        .from('user_profiles')
        .update({ 
          beta_requested_at: new Date().toISOString(),
          beta_notes: 'Requested via dashboard'
        })
        .eq('id', userId)

      if (error) throw error
      
      setRequested(true)
      router.refresh()
    } catch (error) {
      console.error('Error requesting beta access:', error)
      alert('Failed to request beta access. Please try again.')
    } finally {
      setLoading(false)
    }
  }

  const installCellPilot = () => {
    // Open the Apps Script project directly with test deployment
    window.open(
      'https://script.google.com/d/1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O/edit',
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
              Approved on: {new Date(profile.beta_approved_at).toLocaleDateString()}
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
            <summary className="cursor-pointer text-sm text-green-700 hover:text-green-800">
              Installation Instructions
            </summary>
            <div className="mt-3 space-y-2 text-sm text-green-600">
              <p>1. Click "Install CellPilot Now" above</p>
              <p>2. In the Apps Script editor, click Deploy → Test deployments</p>
              <p>3. Click "Install" under Sheets application</p>
              <p>4. Authorize the permissions</p>
              <p>5. Open any Google Sheet and find CellPilot in Extensions menu</p>
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
              Requested on: {new Date(profile.beta_requested_at).toLocaleDateString()}
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