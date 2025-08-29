'use client'

import { ReactNode } from 'react'

interface LayoutWrapperProps {
  children: ReactNode
}

/**
 * Professional layout wrapper that ensures proper spacing
 * and prevents header overlap issues permanently
 */
const LayoutWrapper = ({ children }: LayoutWrapperProps) => {
  return (
    <div className="min-h-screen grid grid-rows-[auto_1fr_auto]">
      {/* Header row - auto height */}
      <div className="sticky top-0 z-header">
        {/* Header component goes here */}
      </div>
      
      {/* Main content row - takes remaining space */}
      <main className="relative">
        {/* All sections go here with natural flow */}
        {children}
      </main>
      
      {/* Footer row - auto height */}
      <footer className="relative">
        {/* Footer component goes here */}
      </footer>
    </div>
  )
}

export default LayoutWrapper