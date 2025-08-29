'use client'

import { useState, useEffect } from 'react'
import Link from 'next/link'

const Hero = () => {
  const [mounted, setMounted] = useState(false)

  useEffect(() => {
    setMounted(true)
  }, [])

  return (
    <section className="snap-section relative overflow-hidden bg-transparent">
      {/* Animated blobs */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="blob-shape w-96 h-96 bg-pastel-mint top-0 -left-48"></div>
        <div className="blob-shape w-96 h-96 bg-pastel-lavender top-48 -right-48 animation-delay-2000"></div>
        <div className="blob-shape w-96 h-96 bg-pastel-sky bottom-0 left-1/2 animation-delay-4000"></div>
      </div>

      {/* Content wrapper with proper top padding to avoid header */}
      <div className="container-wrapper relative z-10 pt-20 sm:pt-24 lg:pt-28 pb-12 sm:pb-16 lg:pb-20">
        <div className="w-full">
          
          {/* Redesigned Hero Layout */}
          <div className={`space-y-6 sm:space-y-8 transition-all duration-700 delay-100 ${mounted ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'}`}>
            {/* Main Title Group */}
            <div className="space-y-2">
              <h2 className="text-2xl sm:text-3xl md:text-4xl lg:text-5xl font-bold text-neutral-800 text-center">
                Google Sheets
              </h2>
              <h1 className="text-center">
                <div className="flex gap-1.5 sm:gap-2 md:gap-3 items-center justify-center flex-wrap max-w-full">
                  {['R', 'e', 'i', 'm', 'a', 'g', 'i', 'n', 'e', 'd'].map((letter, index) => (
                    <span
                      key={index}
                      className="inline-block px-1.5 sm:px-2 md:px-3 lg:px-4 py-1 sm:py-1.5 md:py-2 bg-gradient-to-br from-white/95 to-white/85 backdrop-blur-sm border-2 border-primary-300/60 rounded-lg shadow-md hover:shadow-xl hover:scale-110 hover:-translate-y-1 transition-all cursor-default"
                    style={{
                      opacity: 0,
                      animation: mounted ? `letterPop 0.6s ease-out ${600 + (index * 100)}ms forwards` : 'none',
                      transition: 'transform 0.3s ease, box-shadow 0.3s ease',
                    }}
                  >
                    <span className="text-gradient font-black text-xl sm:text-2xl md:text-3xl lg:text-4xl">{letter}</span>
                  </span>
                ))}
                </div>
              </h1>
            </div>
            
            {/* Tagline */}
            <p className="text-center text-neutral-600 text-base sm:text-lg md:text-xl max-w-2xl mx-auto font-medium">
              Transform your workflow with intelligent automation
            </p>
          </div>
          
          {/* Feature cards - more compact */}
          <div className={`grid grid-cols-1 md:grid-cols-3 gap-4 mt-6 mb-6 transition-all duration-700 delay-300 ${mounted ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'}`}>
            <div className="glass-card p-2 sm:p-3 rounded-xl flex items-center md:flex-col md:text-center gap-2">
              <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-pastel-sky/30 to-pastel-sky/20 flex items-center justify-center flex-shrink-0">
                <svg className="w-4 h-4 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
                </svg>
              </div>
              <div>
                <h3 className="font-semibold text-neutral-800 text-xs">Lightning Fast</h3>
                <p className="text-xs text-neutral-600">Process thousands of rows in seconds</p>
              </div>
            </div>
            <div className="glass-card p-2 sm:p-3 rounded-xl flex items-center md:flex-col md:text-center gap-2">
              <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 flex items-center justify-center flex-shrink-0">
                <svg className="w-4 h-4 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              </div>
              <div>
                <h3 className="font-semibold text-neutral-800 text-xs">Zero Setup</h3>
                <p className="text-xs text-neutral-600">Works instantly with any Google Sheet</p>
              </div>
            </div>
            <div className="glass-card p-2 sm:p-3 rounded-xl flex items-center md:flex-col md:text-center gap-2">
              <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-pastel-lavender/30 to-pastel-lavender/20 flex items-center justify-center flex-shrink-0">
                <svg className="w-4 h-4 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4" />
                </svg>
              </div>
              <div>
                <h3 className="font-semibold text-neutral-800 text-xs">AI-Powered</h3>
                <p className="text-xs text-neutral-600">Smart suggestions and auto-detection</p>
              </div>
            </div>
          </div>


          {/* Preview Window */}
          <div className={`mt-6 sm:mt-8 mb-8 sm:mb-12 transition-all duration-700 delay-500 ${mounted ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'}`}>
            <div className="relative w-full">
              <div className="bg-white/85 backdrop-blur-md rounded-xl sm:rounded-2xl p-0.5 sm:p-1 animate-float shadow-xl sm:shadow-2xl border border-neutral-400/40">
                <div className="bg-gradient-to-br from-white to-neutral-50/50 rounded-lg sm:rounded-xl p-3 sm:p-5">
                  {/* Spreadsheet toolbar */}
                  <div className="flex items-center gap-2 mb-3 sm:mb-4 pb-2 sm:pb-3 border-b border-neutral-200/50">
                    <div className="flex gap-1">
                      <div className="w-1.5 h-1.5 sm:w-2 sm:h-2 rounded-full bg-red-400"></div>
                      <div className="w-1.5 h-1.5 sm:w-2 sm:h-2 rounded-full bg-yellow-400"></div>
                      <div className="w-1.5 h-1.5 sm:w-2 sm:h-2 rounded-full bg-green-400"></div>
                    </div>
                    <div className="text-[10px] sm:text-xs text-neutral-500 font-medium ml-1 sm:ml-2">Data Operations</div>
                  </div>
                  
                  {/* Spreadsheet preview with features */}
                  <div className="overflow-x-auto overflow-y-hidden rounded-md sm:rounded-lg border sm:border-2 border-neutral-300/70 shadow-sm">
                    {/* Column headers */}
                    <div className="grid grid-cols-7 bg-neutral-100/80 border-b sm:border-b-2 border-neutral-300/60 min-w-[320px]">
                      <div className="h-6 sm:h-8 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r sm:border-r-2 border-neutral-300/60"></div>
                      <div className="h-6 sm:h-8 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r border-neutral-200/60">A</div>
                      <div className="h-6 sm:h-8 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r border-neutral-200/60">B</div>
                      <div className="h-6 sm:h-8 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r border-neutral-200/60">C</div>
                      <div className="h-6 sm:h-8 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r border-neutral-200/60">D</div>
                      <div className="h-6 sm:h-8 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r border-neutral-200/60">E</div>
                      <div className="h-6 sm:h-8 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700">F</div>
                    </div>
                    
                    {/* Row 1 - Headers */}
                    <div className="grid grid-cols-7 border-b border-neutral-300/60">
                      <div className="h-8 sm:h-10 bg-neutral-100/80 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r sm:border-r-2 border-neutral-300/60">1</div>
                      <div className="h-8 sm:h-10 bg-gradient-to-r from-pastel-sky/60 to-pastel-sky/30 flex items-center justify-center text-[10px] sm:text-xs font-bold text-neutral-900 border-r border-neutral-200/60">Product</div>
                      <div className="h-8 sm:h-10 bg-gradient-to-r from-pastel-mint/60 to-pastel-mint/30 flex items-center justify-center text-[10px] sm:text-xs font-bold text-neutral-900 border-r border-neutral-200/60">Category</div>
                      <div className="h-8 sm:h-10 bg-gradient-to-r from-pastel-lavender/60 to-pastel-lavender/30 flex items-center justify-center text-[10px] sm:text-xs font-bold text-neutral-900 border-r border-neutral-200/60">Price</div>
                      <div className="h-8 sm:h-10 bg-gradient-to-r from-pastel-peach/60 to-pastel-peach/30 flex items-center justify-center text-[10px] sm:text-xs font-bold text-neutral-900 border-r border-neutral-200/60">Stock</div>
                      <div className="h-8 sm:h-10 bg-gradient-to-r from-pastel-sky/60 to-pastel-sky/30 flex items-center justify-center text-[10px] sm:text-xs font-bold text-neutral-900 border-r border-neutral-200/60">Status</div>
                      <div className="h-8 sm:h-10 bg-gradient-to-r from-pastel-mint/60 to-pastel-mint/30 flex items-center justify-center text-[10px] sm:text-xs font-bold text-neutral-900">Action</div>
                    </div>
                    
                    {/* Row 2 - Data with features */}
                    <div className="grid grid-cols-7 border-b border-neutral-200/60">
                      <div className="h-8 sm:h-10 bg-neutral-100/80 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r sm:border-r-2 border-neutral-300/60">2</div>
                      <div className="h-8 sm:h-10 bg-white flex items-center px-1 sm:px-2 text-[10px] sm:text-xs text-neutral-800 border-r border-neutral-200/60">
                        <span className="text-primary-600 font-semibold">Clean Text</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center px-1 sm:px-2 text-[10px] sm:text-xs text-neutral-800 border-r border-neutral-200/60">
                        <span className="text-purple-600 font-semibold">Format</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center px-1 sm:px-2 text-[10px] sm:text-xs text-neutral-800 border-r border-neutral-200/60">
                        <span className="text-blue-600 font-mono font-semibold">=AVG(D:D)</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center justify-center border-r border-neutral-200/60">
                        <span className="text-orange-600 font-semibold text-[10px] sm:text-xs">Validate</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center justify-center border-r border-neutral-200/60">
                        <span className="px-1 sm:px-2 py-0.5 bg-green-100 text-green-700 rounded font-medium text-[10px] sm:text-xs">Active</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center justify-center text-[10px] sm:text-xs">
                        <span className="text-primary-600 font-semibold">Remove Dupes</span>
                      </div>
                    </div>
                    
                    {/* Row 3 */}
                    <div className="grid grid-cols-7 border-b border-neutral-200/60">
                      <div className="h-8 sm:h-10 bg-neutral-100/80 flex items-center justify-center text-[10px] sm:text-xs font-semibold text-neutral-700 border-r sm:border-r-2 border-neutral-300/60">3</div>
                      <div className="h-8 sm:h-10 bg-white flex items-center px-1 sm:px-2 text-[10px] sm:text-xs text-neutral-800 border-r border-neutral-200/60">
                        <span className="text-teal-600 font-semibold">Trim Spaces</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center px-1 sm:px-2 text-[10px] sm:text-xs text-neutral-800 border-r border-neutral-200/60">
                        <span className="text-indigo-600 font-semibold">Sort A-Z</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center px-1 sm:px-2 text-[10px] sm:text-xs text-neutral-800 border-r border-neutral-200/60">
                        <span className="text-blue-600 font-mono font-semibold">=SUM(D2:D10)</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center justify-center border-r border-neutral-200/60">
                        <span className="text-green-600 font-semibold text-[10px] sm:text-xs">Auto-backup</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center justify-center border-r border-neutral-200/60">
                        <span className="px-1 sm:px-2 py-0.5 bg-yellow-100 text-yellow-700 rounded font-medium text-[10px] sm:text-xs">Pending</span>
                      </div>
                      <div className="h-8 sm:h-10 bg-white flex items-center justify-center text-[10px] sm:text-xs">
                        <span className="text-accent-teal font-semibold">Email Alert</span>
                      </div>
                    </div>
                  </div>
                  
                  {/* Status bar */}
                  <div className="flex items-center justify-between mt-2 sm:mt-3 text-[10px] sm:text-xs text-neutral-500">
                    <span>4 operations available</span>
                    <span className="flex items-center gap-1">
                      <span className="w-1.5 h-1.5 sm:w-2 sm:h-2 rounded-full bg-green-500 animate-pulse"></span>
                      Auto-saved
                    </span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  )
}

export default Hero