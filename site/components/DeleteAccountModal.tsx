'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import { useRouter } from 'next/navigation'

interface DeleteAccountModalProps {
  isOpen: boolean
  onClose: () => void
  userEmail: string
}

export default function DeleteAccountModal({ isOpen, onClose, userEmail }: DeleteAccountModalProps) {
  const [confirmEmail, setConfirmEmail] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const router = useRouter()
  const supabase = createClient()

  const handleDeleteAccount = async () => {
    if (confirmEmail !== userEmail) {
      setError('Email addresses do not match')
      return
    }

    setLoading(true)
    setError(null)

    try {
      // Delete user profile first (due to foreign key constraint)
      const { data: { user } } = await supabase.auth.getUser()
      
      if (!user) {
        throw new Error('No authenticated user found')
      }

      // Delete from profiles table
      const { error: profileError } = await supabase
        .from('profiles')
        .delete()
        .eq('id', user.id)

      if (profileError) {
        throw profileError
      }

      // Delete the user account
      const { error: deleteError } = await supabase.rpc('delete_user')

      if (deleteError) {
        // If custom function doesn't exist, try signing out
        // Note: Full account deletion requires a server-side function
        throw new Error('Account deletion requires contacting support. We\'ve logged you out for now.')
      }

      // Sign out and redirect
      await supabase.auth.signOut()
      router.push('/')
    } catch (err: any) {
      setError(err.message)
      // If error mentions support, still sign out
      if (err.message.includes('support')) {
        setTimeout(async () => {
          await supabase.auth.signOut()
          router.push('/')
        }, 3000)
      }
    } finally {
      setLoading(false)
    }
  }

  if (!isOpen) return null

  return (
    <div className="fixed inset-0 z-modal bg-black/50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl">
        <h2 className="text-xl font-bold text-neutral-900 mb-4">Delete Account</h2>
        
        <div className="bg-red-50 border border-red-200 rounded-lg p-4 mb-6">
          <p className="text-sm text-red-800 font-medium mb-2">Warning: This action cannot be undone</p>
          <p className="text-sm text-red-700">
            Deleting your account will permanently remove all your data, including:
          </p>
          <ul className="list-disc list-inside text-sm text-red-700 mt-2">
            <li>Your profile information</li>
            <li>Installation history</li>
            <li>Usage statistics</li>
            <li>All associated data</li>
          </ul>
        </div>

        {error && (
          <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">
            {error}
          </div>
        )}

        <div className="mb-6">
          <label className="block text-sm font-medium text-neutral-700 mb-2">
            Type your email address to confirm: <strong>{userEmail}</strong>
          </label>
          <input
            type="email"
            value={confirmEmail}
            onChange={(e) => setConfirmEmail(e.target.value)}
            className="w-full px-4 py-2 border border-neutral-300 rounded-lg focus:ring-2 focus:ring-red-500 focus:border-transparent"
            placeholder="Enter your email to confirm"
          />
        </div>

        <div className="flex gap-3">
          <button
            onClick={onClose}
            disabled={loading}
            className="flex-1 px-4 py-2 border border-neutral-300 text-neutral-700 rounded-lg hover:bg-neutral-50 font-medium transition-colors disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleDeleteAccount}
            disabled={loading || confirmEmail !== userEmail}
            className="flex-1 px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {loading ? 'Deleting...' : 'Delete Account'}
          </button>
        </div>
      </div>
    </div>
  )
}