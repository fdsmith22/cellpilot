'use client'

import { useState } from 'react'
import ScrollObserver from './ScrollObserver'

const InteractiveDemo = () => {
  const [activeFeature, setActiveFeature] = useState(0)
  const [hoveredCell, setHoveredCell] = useState<string | null>(null)

  const features = [
    {
      title: 'Tableize Data',
      description: 'Transform messy data into clean, structured columns automatically',
      demo: (
        <div className="space-y-3">
          <div className="bg-white/95 border border-neutral-200 p-3 rounded-lg">
            <div className="text-xs text-neutral-500 mb-1">Messy data:</div>
            <div className="text-sm font-mono text-neutral-700">
              <div>John Doe | Sales | $5000</div>
              <div>Jane Smith | Marketing | $4500</div>
            </div>
          </div>
          <div className="flex items-center justify-center">
            <svg className="w-6 h-6 text-primary-500 animate-bounce" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 14l-7 7m0 0l-7-7m7 7V3" />
            </svg>
          </div>
          <div className="bg-white/95 border border-neutral-200 p-3 rounded-lg bg-gradient-to-br from-pastel-mint/30 to-white/95">
            <div className="text-xs text-neutral-500 mb-1">Structured table:</div>
            <div className="grid grid-cols-3 gap-2 text-xs font-mono">
              <div className="font-bold">Name</div>
              <div className="font-bold">Department</div>
              <div className="font-bold">Salary</div>
              <div>John Doe</div>
              <div>Sales</div>
              <div>$5000</div>
            </div>
          </div>
        </div>
      ),
      icon: null
    },
    {
      title: 'Smart Formula Assistant',
      description: 'Build complex formulas with AI-powered suggestions and validation',
      demo: (
        <div className="space-y-2">
          <div className="bg-white/95 border border-neutral-200 p-3 rounded-lg">
            <div className="text-xs text-neutral-500 mb-1">Type your request:</div>
            <div className="text-sm text-neutral-700">"Sum all sales from Q3 where region is North"</div>
          </div>
          <div className="bg-white/95 border border-neutral-200 p-3 rounded-lg bg-gradient-to-br from-pastel-sky/30 to-white/95">
            <div className="text-xs text-neutral-500 mb-1">Generated formula:</div>
            <code className="text-sm font-mono text-primary-700">=SUMIFS(Sales,Quarter,"Q3",Region,"North")</code>
            <div className="mt-2 text-xs text-green-600">✓ Formula validated and ready to use</div>
          </div>
        </div>
      ),
      icon: null
    },
    {
      title: 'Advanced Restructuring',
      description: 'Reshape, pivot, and transform data with powerful ML-driven tools',
      demo: (
        <div className="relative h-48">
          <div className="grid grid-cols-2 gap-4 h-full">
            <div className="bg-white/95 border border-neutral-200 p-3 rounded-lg">
              <div className="text-xs text-neutral-500 mb-2">Original Format</div>
              <div className="space-y-1 text-xs">
                <div className="bg-neutral-100 p-1 rounded">Date | Product | Sales</div>
                <div className="bg-neutral-50 p-1 rounded">Jan | A | 100</div>
                <div className="bg-neutral-50 p-1 rounded">Jan | B | 150</div>
                <div className="bg-neutral-50 p-1 rounded">Feb | A | 120</div>
              </div>
            </div>
            <div className="bg-gradient-to-br from-pastel-lavender/30 to-white/95 border border-neutral-200 p-3 rounded-lg">
              <div className="text-xs text-neutral-500 mb-2">Pivoted View</div>
              <div className="space-y-1 text-xs">
                <div className="bg-primary-100 p-1 rounded font-semibold">Product | Jan | Feb</div>
                <div className="bg-primary-50 p-1 rounded">A | 100 | 120</div>
                <div className="bg-primary-50 p-1 rounded">B | 150 | -</div>
              </div>
            </div>
          </div>
          <div className="absolute bottom-2 left-1/2 -translate-x-1/2">
            <span className="px-3 py-1 bg-primary-100 text-primary-700 rounded-full text-xs font-medium">
              ML-Powered Transformation
            </span>
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

      </div>
    </section>
  )
}

export default InteractiveDemo