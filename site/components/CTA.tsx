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
      
      <div className="container-wrapper relative z-10 h-full flex items-center justify-center">
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
          
          {/* Trust badges */}
          <div className="flex flex-wrap items-center justify-center gap-6 mt-8 mb-6">
            <div className="flex items-center space-x-2 glass-card px-4 py-2 rounded-full">
              <svg className="w-5 h-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M2.166 4.999A11.954 11.954 0 0010 1.944 11.954 11.954 0 0017.834 5c.11.65.166 1.32.166 2.001 0 5.225-3.34 9.67-8 11.317C5.34 16.67 2 12.225 2 7c0-.682.057-1.35.166-2.001zm11.541 3.708a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              <span className="text-sm font-medium text-neutral-700">Google Verified</span>
            </div>
            <div className="flex items-center space-x-2 glass-card px-4 py-2 rounded-full">
              <svg className="w-5 h-5 text-blue-500" fill="currentColor" viewBox="0 0 20 20">
                <path d="M10.394 2.08a1 1 0 00-.788 0l-7 3a1 1 0 000 1.84L5.25 8.051a.999.999 0 01.356-.257l4-1.714a1 1 0 11.788 1.838L7.667 9.088l1.94.831a1 1 0 00.787 0l7-3a1 1 0 000-1.838l-7-3zM3.31 9.397L5 10.12v4.102a8.969 8.969 0 00-1.05-.174 1 1 0 01-.89-.89 11.115 11.115 0 01.25-3.762zM9.3 16.573A9.026 9.026 0 007 14.935v-3.957l1.818.78a3 3 0 002.364 0l5.508-2.361a11.026 11.026 0 01.25 3.762 1 1 0 01-.89.89 8.968 8.968 0 00-5.35 2.524 1 1 0 01-1.4 0zM6 18a1 1 0 001-1v-2.065a8.935 8.935 0 00-2-.712V17a1 1 0 001 1z"/>
              </svg>
              <span className="text-sm font-medium text-neutral-700">Enterprise Ready</span>
            </div>
          </div>

          <div className="mt-8 sm:mt-12 flex flex-col sm:flex-row items-center justify-center gap-4 sm:gap-6 text-xs sm:text-sm text-neutral-600 px-4">
            <div className="flex items-center">
              <svg className="w-4 h-4 text-primary-500 mr-2" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              <span>Works with all Google accounts</span>
            </div>
            <div className="flex items-center">
              <svg className="w-4 h-4 text-primary-500 mr-2" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              <span>No installation required</span>
            </div>
            <div className="flex items-center">
              <svg className="w-4 h-4 text-primary-500 mr-2" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              <span>25 free operations monthly</span>
            </div>
          </div>
        </ScrollObserver>
      </div>
    </section>
  )
}

export default CTA