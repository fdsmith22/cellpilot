'use client'

import { useState } from 'react'
import Link from 'next/link'
import ScrollObserver from './ScrollObserver'

const CTA = () => {
  const [hoveredStat, setHoveredStat] = useState<number | null>(null)
  return (
    <section className="snap-section relative overflow-hidden bg-transparent">
      {/* Enhanced animated background */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute top-1/4 -left-32 w-96 h-96 bg-gradient-to-br from-pastel-sky/30 to-pastel-lavender/20 rounded-full blur-3xl animate-blob"></div>
        <div className="absolute bottom-1/4 -right-32 w-96 h-96 bg-gradient-to-br from-pastel-mint/30 to-pastel-peach/20 rounded-full blur-3xl animate-blob animation-delay-4000"></div>
        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[600px] h-[600px] bg-gradient-to-r from-primary-100/10 via-accent-teal/5 to-primary-100/10 rounded-full blur-3xl animate-float"></div>
      </div>
      
      {/* Dynamic gradient transition */}
      <div className="absolute inset-0 bg-gradient-to-b from-white/60 via-primary-50/10 to-white/80"></div>
      
      {/* Removed grid pattern - using global grid */}
      
      <div className="container-wrapper relative z-10 py-16 sm:py-20 lg:py-24">
        <ScrollObserver className="max-w-4xl mx-auto text-center px-4">
          <h2 className="text-3xl sm:text-4xl md:text-5xl lg:text-6xl font-bold mb-4 sm:mb-6">
            <span className="text-gradient">Transform</span>
            <span className="text-neutral-900"> Your Spreadsheets </span>
            <span className="text-gradient">Today</span>
          </h2>
          
          <p className="text-base sm:text-lg md:text-xl text-neutral-600 mb-8 sm:mb-10 leading-relaxed max-w-2xl mx-auto">
            Save hours every week with intelligent automation tools designed for Google Sheets power users.
          </p>
          
          <div className="flex flex-col sm:flex-row gap-4 justify-center mb-6">
            <Link
              href="/install"
              className="group relative btn-primary px-8 py-4 text-base font-semibold shadow-lg hover:shadow-xl"
            >
              <span className="relative z-10">Start Free Trial</span>
              <svg className="w-5 h-5 ml-2 inline-block group-hover:translate-x-1 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7l5 5m0 0l-5 5m5-5H6" />
              </svg>
              <div className="absolute inset-0 bg-gradient-to-r from-primary-600 to-primary-500 rounded-full opacity-0 group-hover:opacity-100 transition-opacity"></div>
            </Link>
            <Link
              href="#pricing"
              className="group btn-secondary px-8 py-4 text-base font-semibold"
            >
              <span>View Pricing</span>
              <svg className="w-4 h-4 ml-2 inline-block opacity-0 group-hover:opacity-100 transition-all" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
              </svg>
            </Link>
          </div>
          

        </ScrollObserver>
      </div>
    </section>
  )
}

export default CTA