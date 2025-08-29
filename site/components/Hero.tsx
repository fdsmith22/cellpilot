'use client'

import { useState, useEffect } from 'react'
import Link from 'next/link'

const Hero = () => {
  const [mounted, setMounted] = useState(false)
  const [featureIndex, setFeatureIndex] = useState(0)

  const features = [
    {
      icon: (
        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
        </svg>
      ),
      text: 'Fast & Simple',
      color: 'pastel-mint'
    },
    {
      icon: (
        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
        </svg>
      ),
      text: 'No Setup Required',
      color: 'pastel-lavender'
    },
    {
      icon: (
        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4" />
        </svg>
      ),
      text: 'Smart Tools',
      color: 'pastel-sky'
    }
  ]

  useEffect(() => {
    setMounted(true)
    const interval = setInterval(() => {
      setFeatureIndex((prev) => (prev + 1) % features.length)
    }, 3500) // Slowed down from 2500ms to 3500ms for smoother transitions
    return () => clearInterval(interval)
  }, [])

  return (
    <section className="snap-section relative overflow-hidden bg-transparent">
      {/* Animated blobs */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="blob-shape w-96 h-96 bg-pastel-mint top-0 -left-48"></div>
        <div className="blob-shape w-96 h-96 bg-pastel-lavender top-48 -right-48 animation-delay-2000"></div>
        <div className="blob-shape w-96 h-96 bg-pastel-sky bottom-0 left-1/2 animation-delay-4000"></div>
      </div>

      {/* Clean content wrapper - no complex positioning needed */}
      <div className="container-wrapper relative z-10 py-12 sm:py-16 lg:py-20">
        <div className="max-w-5xl mx-auto">
          
          {/* Heading with new layout */}
          <div className={`mb-8 sm:mb-10 transition-all duration-700 delay-100 ${mounted ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'}`}>
            <div className="text-sm sm:text-base uppercase tracking-[0.3em] text-primary-600 font-semibold text-center mb-4">Google Sheets</div>
            <h1 className="text-5xl sm:text-6xl md:text-7xl lg:text-8xl xl:text-9xl font-bold text-center mb-6">
              <div className="flex gap-2 sm:gap-3 md:gap-4 items-center justify-center flex-wrap max-w-full px-2">
                {['R', 'e', 'i', 'm', 'a', 'g', 'i', 'n', 'e', 'd'].map((letter, index) => (
                  <span
                    key={index}
                    className="inline-block px-3 sm:px-4 md:px-5 lg:px-6 py-2 sm:py-3 md:py-4 bg-gradient-to-br from-white/90 to-white/70 backdrop-blur-sm border-2 border-primary-300/60 rounded-xl shadow-lg hover:shadow-2xl hover:scale-110 hover:-translate-y-2 transition-all cursor-default"
                  style={{
                    opacity: 0,
                    animation: mounted ? `letterPop 0.6s ease-out ${800 + (index * 150)}ms forwards` : 'none',
                    transition: 'transform 0.3s ease, box-shadow 0.3s ease',
                  }}
                >
                  <span className="text-gradient font-black text-3xl sm:text-4xl md:text-5xl lg:text-6xl xl:text-7xl">{letter}</span>
                </span>
              ))}
              </div>
            </h1>
            <div className="text-center text-neutral-600 text-lg sm:text-xl font-medium">Transform your workflow with intelligent automation</div>
          </div>
          
          {/* Feature cards */}
          <div className={`grid grid-cols-1 sm:grid-cols-3 gap-4 mb-8 max-w-4xl mx-auto px-4 transition-all duration-700 delay-200 ${mounted ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'}`}>
            <div className="glass-card p-4 rounded-xl text-center">
              <div className="w-12 h-12 mx-auto mb-2 rounded-lg bg-gradient-to-br from-pastel-sky/30 to-pastel-sky/20 flex items-center justify-center">
                <svg className="w-6 h-6 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
                </svg>
              </div>
              <h3 className="font-semibold text-neutral-800 mb-1">Lightning Fast</h3>
              <p className="text-sm text-neutral-600">Process thousands of rows in seconds</p>
            </div>
            <div className="glass-card p-4 rounded-xl text-center">
              <div className="w-12 h-12 mx-auto mb-2 rounded-lg bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 flex items-center justify-center">
                <svg className="w-6 h-6 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              </div>
              <h3 className="font-semibold text-neutral-800 mb-1">Zero Setup</h3>
              <p className="text-sm text-neutral-600">Works instantly with any Google Sheet</p>
            </div>
            <div className="glass-card p-4 rounded-xl text-center">
              <div className="w-12 h-12 mx-auto mb-2 rounded-lg bg-gradient-to-br from-pastel-lavender/30 to-pastel-lavender/20 flex items-center justify-center">
                <svg className="w-6 h-6 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4" />
                </svg>
              </div>
              <h3 className="font-semibold text-neutral-800 mb-1">AI-Powered</h3>
              <p className="text-sm text-neutral-600">Smart suggestions and auto-detection</p>
            </div>
          </div>

          {/* CTA Buttons as Spreadsheet Cells */}
          <div className={`flex flex-col sm:flex-row gap-4 sm:gap-6 justify-center mb-8 sm:mb-10 px-4 sm:px-0 transition-all duration-700 delay-300 ${mounted ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'}`}>
            {/* Start Free Cell */}
            <div className="relative">
              <div className="absolute -top-3 -left-3 text-xs text-neutral-500 font-mono">B2</div>
              <Link href="/install" className="block">
                <div className="w-full sm:w-48 h-14 bg-white/80 backdrop-blur-md border-2 border-primary-300/50 rounded-lg shadow-lg overflow-hidden relative hover:shadow-xl hover:scale-105 transition-all duration-300 cursor-pointer group">
                  <div className="absolute inset-0 bg-gradient-to-br from-white/50 to-transparent"></div>
                  <div className="relative h-full flex items-center justify-center px-4">
                    <span className="text-sm font-semibold text-primary-700 group-hover:text-primary-800">Start Free</span>
                    <svg className="w-4 h-4 ml-2 text-primary-600 group-hover:translate-x-1 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 8l4 4m0 0l-4 4m4-4H3" />
                    </svg>
                  </div>
                </div>
              </Link>
            </div>

            {/* See How It Works Cell */}
            <div className="relative">
              <div className="absolute -top-3 -left-3 text-xs text-neutral-500 font-mono">C2</div>
              <Link href="#features" className="block">
                <div className="w-full sm:w-48 h-14 bg-white/80 backdrop-blur-md border-2 border-primary-300/50 rounded-lg shadow-lg overflow-hidden relative hover:shadow-xl hover:scale-105 transition-all duration-300 cursor-pointer group">
                  <div className="absolute inset-0 bg-gradient-to-br from-white/50 to-transparent"></div>
                  <div className="relative h-full flex items-center justify-center px-4">
                    <svg className="w-4 h-4 mr-2 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M14.752 11.168l-3.197-2.132A1 1 0 0010 9.87v4.263a1 1 0 001.555.832l3.197-2.132a1 1 0 000-1.664z" />
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                    <span className="text-sm font-semibold text-primary-700 group-hover:text-primary-800">See How It Works</span>
                  </div>
                </div>
              </Link>
            </div>

            {/* Feature Roller Cell */}
            <div className="relative">
              <div className="absolute -top-3 -left-3 text-xs text-neutral-500 font-mono">D2</div>
              <div className="w-full sm:w-48 h-14 bg-white/80 backdrop-blur-md border-2 border-primary-300/50 rounded-lg shadow-lg overflow-hidden relative">
                <div className="absolute inset-0 bg-gradient-to-br from-white/50 to-transparent"></div>
                <div className="relative h-full">
                  {features.map((feature, index) => (
                    <div
                      key={index}
                      className={`absolute inset-0 flex items-center px-4 transition-all duration-700 ease-in-out ${
                        index === featureIndex 
                          ? 'opacity-100 transform translate-y-0' 
                          : index === (featureIndex - 1 + features.length) % features.length
                          ? 'opacity-0 transform -translate-y-full'
                          : 'opacity-0 transform translate-y-full'
                      }`}
                    >
                      <div className={`p-2 rounded-lg ${
                        feature.color === 'pastel-mint' ? 'bg-pastel-mint/20' :
                        feature.color === 'pastel-lavender' ? 'bg-pastel-lavender/20' :
                        'bg-pastel-sky/20'
                      } text-primary-700 mr-3`}>
                        {feature.icon}
                      </div>
                      <span className="text-sm font-semibold text-neutral-800">{feature.text}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>

          {/* Preview Window */}
          <div className={`mt-8 sm:mt-12 mb-12 sm:mb-16 transition-all duration-700 delay-500 ${mounted ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'}`}>
            <div className="relative mx-auto max-w-6xl px-2 sm:px-4">
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