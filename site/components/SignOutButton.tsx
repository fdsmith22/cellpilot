'use client'

import { useAuth } from '@/components/AuthProvider'

export default function SignOutButton() {
  const { signOut } = useAuth()

  return (
    <button
      onClick={signOut}
      className="px-4 py-2 text-sm font-medium text-neutral-600 hover:text-neutral-900 transition-colors"
    >
      Sign Out
    </button>
  )
}