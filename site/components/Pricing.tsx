'use client'

import { useState } from 'react'
import Link from 'next/link'
import ScrollObserver from './ScrollObserver'

const allFeatures = {
  free: [
    '100 operations per month',
    'Tableize Data (basic)',
    'Remove Duplicates',
    'Clean Data (trim, case)',
    'Basic Formula Assistant',
  ],
  starter: [
    '500 operations per month',
    'Everything in Free, plus:',
    'Advanced Tableize with ML',
    'Smart Formula Builder',
    'Cross-Sheet Formulas',
    'Data Validation Generator',
    'Pivot Table Assistant',
    'Usage analytics',
  ],
  professional: [
    'Unlimited operations',
    'Everything in Starter, plus:',
    'Advanced Restructuring',
    'Industry Templates (6 types)',
    'Data Pipeline Manager',
    'Excel Migration Tools',
    'API Integration',
    'Automation Workflows',
    'Batch Operations',
  ],
}

const plans = [
  {
    id: 'free',
    name: 'Free',
    price: '£0',
    period: 'forever',
    description: 'Get started with basic features',
    highlights: [
      '100 operations/month',
      'Basic cleaning tools',
      'Formula assistant',
    ],
    features: allFeatures.free,
    cta: 'Get Started',
    ctaLink: '/install',
    popular: false,
  },
  {
    id: 'starter',
    name: 'Starter',
    price: '£5.99',
    period: 'per month',
    description: 'For regular spreadsheet users',
    highlights: [
      '500 operations/month',
      'ML-powered tools',
      'Advanced formulas',
    ],
    features: allFeatures.starter,
    cta: 'Try Free for 7 Days',
    ctaLink: '/install?plan=starter',
    popular: true,
  },
  {
    id: 'professional',
    name: 'Professional',
    price: '£11.99',
    period: 'per month',
    description: 'For power users and teams',
    highlights: [
      'Unlimited operations',
      'All features included',
      'Industry templates',
    ],
    features: allFeatures.professional,
    cta: 'Try Free for 7 Days',
    ctaLink: '/install?plan=professional',
    popular: false,
  },
]

const Pricing = () => {
  const [expandedPlan, setExpandedPlan] = useState<string | null>(null)
  const [isAnimating, setIsAnimating] = useState(false)
  
  const handlePlanClick = (planId: string) => {
    if (isAnimating) return
    
    setIsAnimating(true)
    
    if (expandedPlan === planId) {
      // Closing
      setExpandedPlan(null)
    } else {
      // Opening or switching
      setExpandedPlan(planId)
    }
    
    // Reset animation lock after transition completes
    setTimeout(() => {
      setIsAnimating(false)
    }, 1200)
  }
  
  return (
    <section id="pricing" className="snap-section relative overflow-hidden bg-transparent">
      <style jsx>{`
        .pricing-grid {
          transition: all 1200ms cubic-bezier(0.25, 0.46, 0.45, 0.94);
        }
        
        .pricing-card {
          transition: all 1200ms cubic-bezier(0.25, 0.46, 0.45, 0.94);
          transform-origin: center;
        }
        
        .pricing-card.hiding {
          opacity: 0;
          transform: scale(0.85) translateY(20px);
          filter: blur(10px);
          pointer-events: none;
        }
        
        .pricing-card.showing {
          opacity: 1;
          transform: scale(1) translateY(0);
          filter: blur(0);
        }
        
        .expanded-card {
          opacity: 0;
          animation: expandedFadeIn 1200ms cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
        }
        
        .expanded-card.closing {
          animation: expandedFadeOut 1200ms cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
        }
        
        @keyframes expandedFadeIn {
          0% {
            opacity: 0;
            transform: scale(0.9) translateY(30px);
          }
          100% {
            opacity: 1;
            transform: scale(1) translateY(0);
          }
        }
        
        @keyframes expandedFadeOut {
          0% {
            opacity: 1;
            transform: scale(1) translateY(0);
          }
          100% {
            opacity: 0;
            transform: scale(0.9) translateY(30px);
          }
        }
        
        .left-panel {
          transition: all 1200ms cubic-bezier(0.25, 0.46, 0.45, 0.94);
          transform-origin: left center;
        }
        
        .right-panel {
          transition: all 1200ms cubic-bezier(0.25, 0.46, 0.45, 0.94);
          transform-origin: right center;
          opacity: 0;
          transform: translateX(60px);
        }
        
        .expanded-card:not(.closing) .left-panel {
          animation: slideLeft 1200ms cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
        }
        
        .expanded-card:not(.closing) .right-panel {
          animation: slideRight 1200ms cubic-bezier(0.25, 0.46, 0.45, 0.94) 200ms forwards;
        }
        
        @keyframes slideLeft {
          0% {
            transform: translateX(0);
          }
          100% {
            transform: translateX(-30px);
          }
        }
        
        @keyframes slideRight {
          0% {
            opacity: 0;
            transform: translateX(60px);
          }
          100% {
            opacity: 1;
            transform: translateX(0);
          }
        }
        
        .feature-item {
          opacity: 0;
          transform: translateX(20px);
          animation: featureFadeIn 600ms cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
        }
        
        .feature-item:nth-child(1) { animation-delay: 400ms; }
        .feature-item:nth-child(2) { animation-delay: 500ms; }
        .feature-item:nth-child(3) { animation-delay: 600ms; }
        .feature-item:nth-child(4) { animation-delay: 700ms; }
        .feature-item:nth-child(5) { animation-delay: 800ms; }
        .feature-item:nth-child(6) { animation-delay: 900ms; }
        .feature-item:nth-child(7) { animation-delay: 1000ms; }
        .feature-item:nth-child(8) { animation-delay: 1100ms; }
        .feature-item:nth-child(9) { animation-delay: 1200ms; }
        
        @keyframes featureFadeIn {
          to {
            opacity: 1;
            transform: translateX(0);
          }
        }
        
        .expanded-card.closing .feature-item {
          animation: none;
          opacity: 1;
          transform: translateX(0);
          transition: all 600ms cubic-bezier(0.25, 0.46, 0.45, 0.94);
          opacity: 0;
          transform: translateX(20px);
        }
      `}</style>
      
      {/* Background decoration */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[400px] sm:w-[600px] md:w-[800px] h-[400px] sm:h-[600px] md:h-[800px] bg-gradient-to-r from-pastel-peach/20 to-pastel-lavender/20 rounded-full blur-3xl"></div>
      </div>
      
      {/* Seamless transition */}
      <div className="absolute inset-0 bg-gradient-to-b from-white/60 via-white/70 to-white/80"></div>
      
      <div className="container-wrapper relative z-10 py-16 sm:py-20 lg:py-24">
        <div className="w-full">
          <ScrollObserver className="text-center mb-6 sm:mb-8">
            <h2 className="text-2xl sm:text-3xl md:text-4xl lg:text-5xl font-bold text-neutral-900 mb-3 sm:mb-4">
              Simple Pricing
            </h2>
            <p className="text-sm sm:text-base md:text-lg text-neutral-600 leading-relaxed">
              Start free and upgrade when you need more. Cancel anytime.
            </p>
            <p className="text-xs sm:text-sm text-neutral-500 mt-2">
              Click on a plan to see all features included
            </p>
          </ScrollObserver>

          {/* Expanded view - rendered separately to avoid layout shifts */}
          {expandedPlan && (
            <div className={`mb-8 overflow-visible expanded-card ${!expandedPlan ? 'closing' : ''}`}>
              {plans.filter(p => p.id === expandedPlan).map(plan => (
                <div key={plan.id} className="max-w-5xl mx-auto">
                  <div className="glass-card rounded-2xl overflow-hidden shadow-2xl">
                    <div className="grid lg:grid-cols-2 gap-0">
                      {/* Plan summary */}
                      <div className="left-panel p-6 lg:p-8">
                        {plan.popular && (
                          <div className="inline-block bg-gradient-to-r from-primary-500 to-accent-teal px-3 py-1 rounded-full text-white text-xs font-medium mb-3">
                            Most Popular
                          </div>
                        )}
                        <h3 className="text-2xl font-bold text-neutral-900 mb-2">{plan.name}</h3>
                        <p className="text-sm text-neutral-600 mb-4">{plan.description}</p>
                        
                        <div className="mb-6">
                          <span className="text-4xl font-bold text-neutral-900">{plan.price}</span>
                          <span className="text-sm text-neutral-600 ml-2">/{plan.period}</span>
                        </div>

                        <Link
                          href={plan.ctaLink}
                          className={`block w-full py-3 px-6 rounded-full text-center font-medium transition-transform duration-300 mb-4 hover:scale-105 ${
                            plan.popular
                              ? 'btn-primary'
                              : 'bg-white border border-neutral-300 text-neutral-700 hover:bg-neutral-100'
                          }`}
                        >
                          {plan.cta}
                        </Link>

                        <button
                          onClick={() => handlePlanClick(plan.id)}
                          className="text-sm text-neutral-500 hover:text-neutral-700 transition-colors duration-300"
                        >
                          ← Back to all plans
                        </button>
                      </div>

                      {/* Full feature list */}
                      <div className="right-panel p-6 lg:p-8 bg-gradient-to-br from-neutral-50/50 to-white">
                        <h4 className="text-lg font-semibold text-neutral-900 mb-4">All Features Included:</h4>
                        <ul className="space-y-3">
                          {plan.features.map((feature, index) => (
                            <li key={index} className="feature-item flex items-start">
                              <svg
                                className="h-5 w-5 text-accent-teal mt-0.5 mr-3 flex-shrink-0"
                                fill="currentColor"
                                viewBox="0 0 20 20"
                              >
                                <path
                                  fillRule="evenodd"
                                  d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z"
                                  clipRule="evenodd"
                                />
                              </svg>
                              <span className={`text-sm text-neutral-700 ${
                                feature.startsWith('Everything') ? 'font-semibold text-primary-600' : ''
                              }`}>{feature}</span>
                            </li>
                          ))}
                        </ul>
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* Collapsed view */}
          <div className={`pricing-grid grid grid-cols-1 lg:grid-cols-3 gap-6 lg:gap-8 items-stretch mb-6 sm:mb-8 ${
            expandedPlan ? 'pointer-events-none' : ''
          }`}>
            {plans.map((plan, index) => (
              <ScrollObserver
                key={plan.name}
                className={`scroll-observer-delay-${index * 100}`}
              >
                <div
                  onClick={() => handlePlanClick(plan.id)}
                  className={`pricing-card glass-card rounded-xl sm:rounded-2xl overflow-hidden h-full cursor-pointer hover:scale-105 hover:shadow-2xl ${
                    plan.popular ? 'ring-2 ring-primary-300' : ''
                  } ${
                    expandedPlan ? 'hiding' : 'showing'
                  }`}
                >
                  {plan.popular && (
                    <div className="bg-gradient-to-r from-primary-500 to-accent-teal px-4 py-2 text-center">
                      <span className="text-white text-sm font-medium">Most Popular</span>
                    </div>
                  )}
                  
                  <div className="p-4 sm:p-6">
                    <h3 className="text-lg sm:text-xl font-bold text-neutral-900 mb-1">{plan.name}</h3>
                    <p className="text-xs sm:text-sm text-neutral-600 mb-3">{plan.description}</p>
                    
                    <div className="mb-4">
                      <span className="text-2xl sm:text-3xl font-bold text-neutral-900">{plan.price}</span>
                      <span className="text-xs sm:text-sm text-neutral-600 ml-1">/{plan.period}</span>
                    </div>

                    <ul className="space-y-2 mb-6">
                      {plan.highlights.map((feature, index) => (
                        <li key={index} className="flex items-start">
                          <svg
                            className="h-4 w-4 text-accent-teal mt-0.5 mr-2 flex-shrink-0"
                            fill="currentColor"
                            viewBox="0 0 20 20"
                          >
                            <path
                              fillRule="evenodd"
                              d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z"
                              clipRule="evenodd"
                            />
                          </svg>
                          <span className="text-xs sm:text-sm text-neutral-700">{feature}</span>
                        </li>
                      ))}
                    </ul>

                    <div className="text-center">
                      <span className="text-xs text-primary-600 font-medium">
                        Click to see all {plan.features.length} features →
                      </span>
                    </div>
                  </div>
                </div>
              </ScrollObserver>
            ))}
          </div>

          <div className="mt-8 sm:mt-10 text-center px-4">
            <div className="inline-flex flex-col sm:flex-row items-start sm:items-center gap-4 sm:gap-8 p-4 sm:p-6 rounded-xl sm:rounded-2xl glass-card">
              <div className="flex items-center">
                <svg className="w-4 sm:w-5 h-4 sm:h-5 text-accent-teal mr-2 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                  <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                </svg>
                <span className="text-sm sm:text-base text-neutral-700 font-medium">Cancel anytime</span>
              </div>
              <div className="flex items-center">
                <svg className="w-4 sm:w-5 h-4 sm:h-5 text-accent-teal mr-2 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                  <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                </svg>
                <span className="text-sm sm:text-base text-neutral-700 font-medium">Secure payments</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  )
}

export default Pricing