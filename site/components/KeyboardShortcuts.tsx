'use client'

import { useEffect, useState } from 'react'

export default function KeyboardShortcuts() {
  const [showHelper, setShowHelper] = useState(false)

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Show help with '?'
      if (e.key === '?' && !e.ctrlKey && !e.metaKey) {
        e.preventDefault()
        setShowHelper(prev => !prev)
      }
      
      // Quick navigation
      if (e.ctrlKey || e.metaKey) {
        switch(e.key) {
          case '1':
            e.preventDefault()
            document.querySelector('#features')?.scrollIntoView({ behavior: 'smooth' })
            break
          case '2':
            e.preventDefault()
            document.querySelector('#pricing')?.scrollIntoView({ behavior: 'smooth' })
            break
          case '3':
            e.preventDefault()
            document.querySelector('.snap-section:last-of-type')?.scrollIntoView({ behavior: 'smooth' })
            break
          case 'i':
            e.preventDefault()
            window.location.href = '/install'
            break
        }
      }
      
      // ESC to close helper
      if (e.key === 'Escape') {
        setShowHelper(false)
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [])

  return (
    <>
      {/* Subtle hint on desktop */}
      <div className="hidden lg:block fixed bottom-4 left-4 z-40">
        <button
          onClick={() => setShowHelper(true)}
          className="glass-card px-3 py-2 text-xs text-neutral-600 hover:text-neutral-900 transition-all hover:scale-105"
          aria-label="Keyboard shortcuts"
        >
          <kbd className="font-mono bg-white/50 px-1.5 py-0.5 rounded">?</kbd> Shortcuts
        </button>
      </div>

      {/* Shortcuts modal */}
      {showHelper && (
        <div 
          className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4"
          onClick={() => setShowHelper(false)}
        >
          <div 
            className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-6"
            onClick={e => e.stopPropagation()}
          >
            <h3 className="text-lg font-semibold text-neutral-900 mb-4">Keyboard Shortcuts</h3>
            
            <div className="space-y-3">
              <div className="flex justify-between items-center">
                <span className="text-sm text-neutral-600">Toggle this help</span>
                <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">?</kbd>
              </div>
              
              <div className="flex justify-between items-center">
                <span className="text-sm text-neutral-600">Features section</span>
                <div className="flex gap-1">
                  <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">⌘</kbd>
                  <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">1</kbd>
                </div>
              </div>
              
              <div className="flex justify-between items-center">
                <span className="text-sm text-neutral-600">Pricing section</span>
                <div className="flex gap-1">
                  <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">⌘</kbd>
                  <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">2</kbd>
                </div>
              </div>
              
              <div className="flex justify-between items-center">
                <span className="text-sm text-neutral-600">Get Started</span>
                <div className="flex gap-1">
                  <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">⌘</kbd>
                  <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">3</kbd>
                </div>
              </div>
              
              <div className="flex justify-between items-center">
                <span className="text-sm text-neutral-600">Install page</span>
                <div className="flex gap-1">
                  <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">⌘</kbd>
                  <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">I</kbd>
                </div>
              </div>
              
              <div className="flex justify-between items-center">
                <span className="text-sm text-neutral-600">Close dialog</span>
                <kbd className="font-mono text-xs bg-neutral-100 px-2 py-1 rounded">ESC</kbd>
              </div>
            </div>

            <button
              onClick={() => setShowHelper(false)}
              className="mt-6 w-full btn-primary py-2"
            >
              Got it
            </button>
          </div>
        </div>
      )}
    </>
  )
}