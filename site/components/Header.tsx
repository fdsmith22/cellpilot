'use client'

import { useState, useEffect } from 'react'
import Link from 'next/link'

const Header = () => {
  const [scrolled, setScrolled] = useState(false)
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false)

  useEffect(() => {
    const handleScroll = () => {
      setScrolled(window.scrollY > 20)
    }
    window.addEventListener('scroll', handleScroll)
    return () => window.removeEventListener('scroll', handleScroll)
  }, [])

  const navItems = [
    { name: 'Features', href: '/#features' },
    { name: 'Pricing', href: '/#pricing' },
    { name: 'Documentation', href: '/docs' },
  ]

  return (
    <header className={`fixed top-0 left-0 right-0 z-50 transition-all duration-300 ${
      scrolled 
        ? 'bg-gradient-to-b from-white/95 via-white/85 to-transparent backdrop-blur-xl border-b border-neutral-200/20' 
        : 'bg-gradient-to-b from-white/80 via-white/60 to-transparent backdrop-blur-lg'
    }`}>
      {/* Grid pattern that fades into header */}
      <div className="absolute inset-0 grid-pattern opacity-10 pointer-events-none" 
           style={{maskImage: 'linear-gradient(to bottom, black, transparent)', WebkitMaskImage: 'linear-gradient(to bottom, black, transparent)'}}></div>
      <nav className="relative py-3 max-sm:py-2">
        <div className="flex items-center justify-between min-h-[72px] max-sm:min-h-[60px]">
          {/* Logo - positioned in top left */}
          <Link href="/" className="flex items-center pl-4 md:pl-6 lg:pl-8">
            <img 
              src="/logo/combined/horizontal-standard-200x60.svg" 
              alt="CellPilot" 
              className="h-20 lg:h-16 md:h-14 sm:h-12 max-sm:h-10 w-auto"
            />
          </Link>

          {/* Desktop Navigation */}
          <div className="hidden md:flex items-center gap-3 lg:gap-4 pr-4 md:pr-6 lg:pr-8">
            {navItems.map((item, index) => {
              const colors = [
                'from-pastel-sky/30 to-pastel-sky/20 border-pastel-sky/40',
                'from-pastel-mint/30 to-pastel-mint/20 border-pastel-mint/40',
                'from-pastel-lavender/30 to-pastel-lavender/20 border-pastel-lavender/40'
              ]
              return (
                <div key={item.name} className="relative">
                  <div className="absolute -top-2 -left-2 text-[10px] text-neutral-500/70 font-mono">
                    {String.fromCharCode(65 + index)}{index + 1}
                  </div>
                  <a
                    href={item.href}
                    className={`block px-4 py-2 bg-gradient-to-br ${colors[index]} backdrop-blur-md border rounded-lg hover:shadow-lg hover:scale-105 transition-all text-sm font-medium text-neutral-700 hover:text-primary-700`}
                    onClick={(e) => {
                      if (item.href.startsWith('/#')) {
                        // If on home page, just scroll
                        if (window.location.pathname === '/') {
                          e.preventDefault()
                          const element = document.querySelector(item.href.substring(1))
                          element?.scrollIntoView({ behavior: 'smooth' })
                        }
                        // Otherwise, let the browser navigate to home page with hash
                      }
                    }}
                  >
                    {item.name}
                  </a>
                </div>
              )
            })}
            
            {/* CTA Button as Cell */}
            <div className="relative ml-2">
              <div className="absolute -top-2 -left-2 text-[10px] text-neutral-500/70 font-mono">D1</div>
              <Link 
                href="/install" 
                className="block px-5 py-2 bg-gradient-to-br from-primary-500/95 to-primary-600/95 text-white rounded-lg border border-primary-600/50 hover:from-primary-600 hover:to-primary-700 hover:shadow-xl hover:scale-105 transition-all text-sm font-medium shadow-md backdrop-blur-sm"
              >
                Get Started
              </Link>
            </div>
          </div>

          {/* Mobile Menu Button */}
          <button
            className="md:hidden p-2 mr-4 rounded-lg hover:bg-neutral-100 transition-colors"
            onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
            aria-label="Toggle navigation menu"
            aria-expanded={mobileMenuOpen}
          >
            <svg className="w-6 h-6 text-neutral-700" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              {mobileMenuOpen ? (
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
              ) : (
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 12h16M4 18h16" />
              )}
            </svg>
          </button>
        </div>

        {/* Mobile Menu */}
        {mobileMenuOpen && (
          <div className="md:hidden mt-4 pb-4 border-t border-neutral-200">
            <div className="pt-4 space-y-3">
              {navItems.map((item, index) => {
                const colors = [
                  'from-pastel-sky/30 to-pastel-sky/20 border-pastel-sky/40',
                  'from-pastel-mint/30 to-pastel-mint/20 border-pastel-mint/40',
                  'from-pastel-lavender/30 to-pastel-lavender/20 border-pastel-lavender/40'
                ]
                return (
                  <div key={item.name} className="relative px-2">
                    <div className="absolute -top-1 left-0 text-[10px] text-neutral-400 font-mono">
                      {String.fromCharCode(65 + index)}{index + 1}
                    </div>
                    <a
                      href={item.href}
                      className={`block px-4 py-3 bg-gradient-to-br ${colors[index]} backdrop-blur-sm border rounded-lg transition-all text-sm font-medium text-neutral-700`}
                      onClick={(e) => {
                        if (item.href.startsWith('/#')) {
                          // If on home page, just scroll
                          if (window.location.pathname === '/') {
                            e.preventDefault()
                            const element = document.querySelector(item.href.substring(1))
                            element?.scrollIntoView({ behavior: 'smooth' })
                          }
                          // Otherwise, let the browser navigate to home page with hash
                        }
                        setMobileMenuOpen(false)
                      }}
                    >
                      {item.name}
                    </a>
                  </div>
                )
              })}
              <div className="relative px-2">
                <div className="absolute -top-1 left-0 text-[10px] text-neutral-400 font-mono">D1</div>
                <Link
                  href="/install"
                  className="block px-4 py-3 bg-gradient-to-br from-primary-500/95 to-primary-600/95 text-white rounded-lg border border-primary-600/50 text-center font-medium shadow-md backdrop-blur-sm"
                  onClick={() => setMobileMenuOpen(false)}
                >
                  Get Started
                </Link>
              </div>
            </div>
          </div>
        )}
      </nav>
    </header>
  )
}

export default Header