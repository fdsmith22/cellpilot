'use client'

import { useState } from 'react'
import DeleteAccountModal from './DeleteAccountModal'

interface DangerZoneProps {
  userEmail: string
}

export default function DangerZone({ userEmail }: DangerZoneProps) {
  const [showDeleteModal, setShowDeleteModal] = useState(false)
  const [isExpanded, setIsExpanded] = useState(false)

  return (
    <>
      <details className="group">
        <summary 
          className="flex items-center justify-between cursor-pointer text-sm font-medium text-red-800 hover:text-red-900"
          onClick={(e) => {
            e.preventDefault()
            setIsExpanded(!isExpanded)
          }}
        >
          <span className="flex items-center gap-2">
            <svg 
              className={`w-4 h-4 transition-transform ${isExpanded ? 'rotate-90' : ''}`} 
              fill="none" 
              stroke="currentColor" 
              viewBox="0 0 24 24"
            >
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
            </svg>
            Advanced Settings
          </span>
        </summary>
        
        {isExpanded && (
          <div className="mt-4 pt-4 border-t border-red-200">
            <div className="flex items-start justify-between">
              <div>
                <h4 className="text-sm font-medium text-red-900">Delete Account</h4>
                <p className="text-xs text-red-700 mt-1">
                  Permanently delete your account and all data
                </p>
              </div>
              <button
                onClick={() => setShowDeleteModal(true)}
                className="px-3 py-1.5 bg-red-600 text-white rounded text-xs hover:bg-red-700 font-medium transition-colors"
              >
                Delete
              </button>
            </div>
          </div>
        )}
      </details>

      <DeleteAccountModal
        isOpen={showDeleteModal}
        onClose={() => setShowDeleteModal(false)}
        userEmail={userEmail}
      />
    </>
  )
}