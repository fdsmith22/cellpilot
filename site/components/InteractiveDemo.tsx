'use client'

import { useState } from 'react'
import ScrollObserver from './ScrollObserver'

const InteractiveDemo = () => {
  const [activeFeature, setActiveFeature] = useState(0)
  const [hoveredCell, setHoveredCell] = useState<string | null>(null)

  const features = [
    {
      title: 'Smart Formula Builder',
      description: 'Describe what you want in plain English and get the perfect formula',
      demo: (
        <div className="space-y-3">
          <div className="bg-white/95 border border-neutral-200 p-3 rounded-lg">
            <div className="text-xs text-neutral-500 mb-1">Your request:</div>
            <div className="text-sm text-neutral-700">"Calculate the average of column B where column A contains 'Active'"</div>
          </div>
          <div className="flex items-center justify-center">
            <svg className="w-6 h-6 text-primary-500 animate-bounce" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 14l-7 7m0 0l-7-7m7 7V3" />
            </svg>
          </div>
          <div className="bg-white/95 border border-neutral-200 p-3 rounded-lg bg-gradient-to-br from-pastel-mint/30 to-white/95">
            <div className="text-xs text-neutral-500 mb-1">Generated formula:</div>
            <code className="text-sm font-mono text-primary-700">=AVERAGEIF(A:A,"Active",B:B)</code>
          </div>
        </div>
      ),
      icon: null
    },
    {
      title: 'Data Cleaning Assistant',
      description: 'Automatically detect and fix common data issues with one click',
      demo: (
        <div className="space-y-2">
          <div className="grid grid-cols-2 gap-2">
            <div className="bg-white/95 border border-neutral-200 p-2 rounded">
              <div className="text-xs text-neutral-500 mb-1">Before:</div>
              <div className="space-y-1 text-xs font-mono">
                <div className="text-red-600">  john doe  </div>
                <div className="text-red-600">JANE SMITH</div>
                <div className="text-red-600">bob_jones@email</div>
              </div>
            </div>
            <div className="bg-white/95 border border-neutral-200 p-2 rounded bg-gradient-to-br from-pastel-sky/30 to-white/95">
              <div className="text-xs text-neutral-500 mb-1">After:</div>
              <div className="space-y-1 text-xs font-mono">
                <div className="text-green-600">John Doe</div>
                <div className="text-green-600">Jane Smith</div>
                <div className="text-green-600">bob_jones@email.com</div>
              </div>
            </div>
          </div>
          <div className="flex justify-center">
            <span className="px-3 py-1 bg-green-100 text-green-700 rounded-full text-xs font-medium">
              ✓ 3 issues fixed automatically
            </span>
          </div>
        </div>
      ),
      icon: null
    },
    {
      title: 'Visual Data Mapper',
      description: 'See relationships between your data with interactive visualizations',
      demo: (
        <div className="relative h-48">
          <div className="absolute top-4 left-4 w-20 h-20 bg-gradient-to-br from-pastel-lavender/40 to-pastel-lavender/20 rounded-full flex items-center justify-center">
            <span className="text-xs font-semibold">Sales</span>
          </div>
          <div className="absolute top-4 right-4 w-20 h-20 bg-gradient-to-br from-pastel-mint/40 to-pastel-mint/20 rounded-full flex items-center justify-center">
            <span className="text-xs font-semibold">Products</span>
          </div>
          <div className="absolute bottom-4 left-1/2 -translate-x-1/2 w-20 h-20 bg-gradient-to-br from-pastel-sky/40 to-pastel-sky/20 rounded-full flex items-center justify-center">
            <span className="text-xs font-semibold">Customers</span>
          </div>
          <svg className="absolute inset-0 w-full h-full" style={{ zIndex: -1 }}>
            <line x1="60" y1="60" x2="150" y2="150" stroke="rgb(168, 162, 255)" strokeWidth="2" strokeDasharray="5,5" opacity="0.5" />
            <line x1="240" y1="60" x2="150" y2="150" stroke="rgb(162, 255, 203)" strokeWidth="2" strokeDasharray="5,5" opacity="0.5" />
            <line x1="60" y1="60" x2="240" y2="60" stroke="rgb(173, 216, 255)" strokeWidth="2" strokeDasharray="5,5" opacity="0.5" />
          </svg>
          <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2">
            <div className="animate-pulse text-primary-600 text-xs font-medium">
              3 relationships detected
            </div>
          </div>
        </div>
      ),
      icon: null
    }
  ]

  return (
    <section className="relative overflow-hidden bg-transparent py-16 sm:py-20">
      {/* Background decoration */}
      <div className="absolute inset-0">
        <div className="absolute top-1/3 left-0 w-96 h-96 bg-pastel-lavender/10 rounded-full blur-3xl"></div>
        <div className="absolute bottom-1/3 right-0 w-96 h-96 bg-pastel-mint/10 rounded-full blur-3xl"></div>
      </div>
      
      <div className="container-wrapper relative z-10">
        <ScrollObserver className="text-center mb-12">
          <h2 className="text-3xl sm:text-4xl md:text-5xl font-bold text-neutral-900 mb-4">
            How It Works
          </h2>
          <p className="text-lg text-neutral-600">
            Simple tools that help you work faster with spreadsheets
          </p>
        </ScrollObserver>

        <div className="w-full">
          <div className="grid lg:grid-cols-2 gap-8 items-center">
            {/* Feature selector */}
            <div className="space-y-4">
              {features.map((feature, index) => (
                <ScrollObserver
                  key={index}
                  className={`scroll-observer-delay-${index * 100}`}
                >
                  <button
                    onClick={() => setActiveFeature(index)}
                    className={`w-full text-left p-6 rounded-2xl transition-all ${
                      activeFeature === index
                        ? 'bg-white border-2 border-primary-300 shadow-xl scale-105'
                        : 'bg-white/90 border border-neutral-200 hover:shadow-lg hover:scale-102'
                    }`}
                  >
                    <div>
                      <h3 className="text-lg font-semibold text-neutral-900 mb-2">
                        {feature.title}
                      </h3>
                      <p className="text-sm text-neutral-600">
                        {feature.description}
                      </p>
                    </div>
                  </button>
                </ScrollObserver>
              ))}
            </div>

            {/* Demo area */}
            <ScrollObserver>
              <div className="bg-white/95 border border-neutral-200 shadow-lg rounded-2xl p-8 min-h-[400px] flex items-center justify-center">
                <div className="w-full">
                  {features[activeFeature].demo}
                </div>
              </div>
            </ScrollObserver>
          </div>
        </div>

        {/* Live spreadsheet interaction */}
        <div className="mt-16">
          <ScrollObserver>
            <div className="bg-white/95 border border-neutral-200 shadow-lg rounded-2xl p-6 max-w-5xl mx-auto">
              <h3 className="text-xl font-semibold text-neutral-900 mb-4 text-center">
                Try It Yourself - Hover Over Cells
              </h3>
              <div className="overflow-hidden rounded-lg border-2 border-neutral-200">
                <div className="grid grid-cols-5 bg-neutral-100">
                  <div className="p-2 text-center text-xs font-semibold border-r border-b border-neutral-200"></div>
                  {['A', 'B', 'C', 'D'].map(col => (
                    <div key={col} className="p-2 text-center text-xs font-semibold border-r border-b border-neutral-200">
                      {col}
                    </div>
                  ))}
                </div>
                {[1, 2, 3].map(row => (
                  <div key={row} className="grid grid-cols-5">
                    <div className="p-2 text-center text-xs font-semibold bg-neutral-100 border-r border-b border-neutral-200">
                      {row}
                    </div>
                    {['A', 'B', 'C', 'D'].map(col => {
                      const cellId = `${col}${row}`
                      return (
                        <div
                          key={cellId}
                          className={`p-3 border-r border-b border-neutral-200 cursor-pointer transition-all ${
                            hoveredCell === cellId
                              ? 'bg-gradient-to-br from-primary-100 to-primary-50 scale-105 shadow-lg z-10 relative'
                              : 'bg-white hover:bg-neutral-50'
                          }`}
                          onMouseEnter={() => setHoveredCell(cellId)}
                          onMouseLeave={() => setHoveredCell(null)}
                        >
                          <div className="text-xs text-center">
                            {hoveredCell === cellId && (
                              <span className="text-primary-600 font-semibold animate-pulse">
                                Click to edit
                              </span>
                            )}
                          </div>
                        </div>
                      )
                    })}
                  </div>
                ))}
              </div>
              {hoveredCell && (
                <div className="mt-3 text-center">
                  <span className="text-sm text-neutral-600">
                    Selected cell: <span className="font-mono font-semibold text-primary-600">{hoveredCell}</span>
                  </span>
                </div>
              )}
            </div>
          </ScrollObserver>
        </div>
      </div>
    </section>
  )
}

export default InteractiveDemo