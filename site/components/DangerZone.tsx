'use client'

import { useState } from 'react'
import DeleteAccountModal from './DeleteAccountModal'

interface DangerZoneProps {
  userEmail: string
}

export default function DangerZone({ userEmail }: DangerZoneProps) {
  const [showDeleteModal, setShowDeleteModal] = useState(false)

  return (
    <>
      <div className="space-y-4">
        <div className="flex items-start justify-between">
          <div>
            <h3 className="text-lg font-medium text-red-900">Delete Account</h3>
            <p className="text-sm text-red-700 mt-1">
              Permanently delete your account and all associated data. This action cannot be undone.
            </p>
          </div>
          <button
            onClick={() => setShowDeleteModal(true)}
            className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 font-medium transition-colors text-sm whitespace-nowrap"
          >
            Delete Account
          </button>
        </div>
      </div>

      <DeleteAccountModal
        isOpen={showDeleteModal}
        onClose={() => setShowDeleteModal(false)}
        userEmail={userEmail}
      />
    </>
  )
}