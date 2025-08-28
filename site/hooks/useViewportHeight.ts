'use client'

import { useEffect } from 'react'

/**
 * Hook to fix viewport height issues on mobile browsers
 * Sets CSS custom property --vh to actual viewport height
 * Handles browser chrome changes on scroll
 */
export function useViewportHeight() {
  useEffect(() => {
    // Set initial viewport height
    const setViewportHeight = () => {
      const vh = window.innerHeight * 0.01
      document.documentElement.style.setProperty('--vh', `${vh}px`)
    }

    // Set on mount
    setViewportHeight()

    // Update on resize and orientation change
    const handleResize = () => {
      setViewportHeight()
    }

    window.addEventListener('resize', handleResize)
    window.addEventListener('orientationchange', handleResize)
    
    // Also update on scroll for mobile browsers that hide/show address bar
    let ticking = false
    const handleScroll = () => {
      if (!ticking) {
        window.requestAnimationFrame(() => {
          setViewportHeight()
          ticking = false
        })
        ticking = true
      }
    }
    
    window.addEventListener('scroll', handleScroll, { passive: true })

    return () => {
      window.removeEventListener('resize', handleResize)
      window.removeEventListener('orientationchange', handleResize)
      window.removeEventListener('scroll', handleScroll)
    }
  }, [])
}

// Hook for getting current viewport dimensions
export function useViewportDimensions() {
  useEffect(() => {
    const updateDimensions = () => {
      const vw = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0)
      const vh = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0)
      
      document.documentElement.style.setProperty('--vw', `${vw}px`)
      document.documentElement.style.setProperty('--vh-real', `${vh}px`)
    }

    updateDimensions()
    window.addEventListener('resize', updateDimensions)
    
    return () => window.removeEventListener('resize', updateDimensions)
  }, [])
}