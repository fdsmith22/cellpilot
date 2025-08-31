'use client'

import { useEffect, useState } from 'react'
import { useSearchParams } from 'next/navigation'

export default function WelcomeMessage() {
  const searchParams = useSearchParams()
  const [show, setShow] = useState(false)

  useEffect(() => {
    if (searchParams.get('welcome') === 'true') {
      setShow(true)
      // Auto-hide after 5 seconds
      const timer = setTimeout(() => setShow(false), 5000)
      
      // Clean up URL
      const url = new URL(window.location.href)
      url.searchParams.delete('welcome')
      window.history.replaceState({}, '', url)
      
      return () => clearTimeout(timer)
    }
  }, [searchParams])

  if (!show) return null

  return (
    <div className="mb-6 p-4 bg-gradient-to-r from-green-50 to-emerald-50 border border-green-200 rounded-lg animate-fade-in">
      <div className="flex items-start">
        <div className="flex-shrink-0">
          <svg className="w-6 h-6 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
        </div>
        <div className="ml-3">
          <h3 className="text-sm font-medium text-green-900">
            Welcome to CellPilot!
          </h3>
          <p className="mt-1 text-sm text-green-700">
            Your email has been verified successfully. Get started by requesting beta access below.
          </p>
        </div>
        <button
          onClick={() => setShow(false)}
          className="ml-auto flex-shrink-0 text-green-500 hover:text-green-600"
        >
          <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
          </svg>
        </button>
      </div>
    </div>
  )
}