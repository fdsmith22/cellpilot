'use client'

import { useEffect, useState } from 'react'

export default function ScrollProgress() {
  const [progress, setProgress] = useState(0)
  const [isVisible, setIsVisible] = useState(false)

  useEffect(() => {
    const updateProgress = () => {
      const scrollHeight = document.documentElement.scrollHeight - document.documentElement.clientHeight
      const scrolled = window.scrollY
      const progressPercentage = (scrolled / scrollHeight) * 100
      setProgress(progressPercentage)
      
      // Show progress bar after scrolling down a bit
      setIsVisible(scrolled > 100)
    }

    updateProgress()
    window.addEventListener('scroll', updateProgress, { passive: true })
    
    return () => window.removeEventListener('scroll', updateProgress)
  }, [])

  return (
    <>
      {/* Top progress bar */}
      <div 
        className={`fixed top-0 left-0 right-0 h-0.5 z-[60] transition-opacity duration-300 ${
          isVisible ? 'opacity-100' : 'opacity-0'
        }`}
      >
        <div 
          className="h-full bg-gradient-to-r from-primary-400 to-accent-teal transition-all duration-150 ease-out"
          style={{ width: `${progress}%` }}
        />
      </div>
      
      {/* Side dot indicator for desktop */}
      <div className="hidden lg:block fixed right-8 top-1/2 -translate-y-1/2 z-40">
        <div className="flex flex-col gap-4">
          {['Hero', 'Features', 'Pricing', 'Get Started'].map((section, index) => {
            const sectionProgress = (index / 3) * 100
            const isActive = progress >= sectionProgress && progress < sectionProgress + 33.33
            
            return (
              <button
                key={section}
                onClick={() => {
                  const element = document.querySelectorAll('.snap-section')[index]
                  element?.scrollIntoView({ behavior: 'smooth' })
                }}
                className="group flex items-center justify-end gap-3"
                aria-label={`Scroll to ${section}`}
              >
                <span className={`text-xs font-medium transition-all duration-300 ${
                  isActive ? 'opacity-100 translate-x-0' : 'opacity-0 translate-x-2'
                } group-hover:opacity-100 group-hover:translate-x-0`}>
                  {section}
                </span>
                <div className={`w-2 h-2 rounded-full transition-all duration-300 ${
                  isActive 
                    ? 'bg-primary-500 scale-125' 
                    : 'bg-neutral-300 hover:bg-primary-400'
                }`} />
              </button>
            )
          })}
        </div>
      </div>
    </>
  )
}