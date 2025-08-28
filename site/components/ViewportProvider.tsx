'use client'

import { useViewportHeight, useViewportDimensions } from '@/hooks/useViewportHeight'
import { useEffect } from 'react'
import { isMobile, isIOS, isAndroid } from 'react-device-detect'

export default function ViewportProvider({ children }: { children: React.ReactNode }) {
  // Initialize viewport height fixes
  useViewportHeight()
  useViewportDimensions()
  
  // Add device-specific classes to html element
  useEffect(() => {
    const html = document.documentElement
    
    // Add device classes
    if (isMobile) html.classList.add('is-mobile')
    if (isIOS) html.classList.add('is-ios')
    if (isAndroid) html.classList.add('is-android')
    
    // Add viewport-fit support check
    if (CSS.supports('padding-top', 'env(safe-area-inset-top)')) {
      html.classList.add('supports-safe-area')
    }
    
    // Detect if device has notch/safe areas
    const hasNotch = () => {
      const safeTop = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--safe-top'))
      return safeTop > 0
    }
    
    if (hasNotch()) {
      html.classList.add('has-notch')
    }
    
    return () => {
      html.classList.remove('is-mobile', 'is-ios', 'is-android', 'supports-safe-area', 'has-notch')
    }
  }, [])
  
  return <>{children}</>
}