'use client'

import { useState, useEffect } from 'react'
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
  const [isTransitioning, setIsTransitioning] = useState(false)
  const [animationPhase, setAnimationPhase] = useState<'idle' | 'collapsing' | 'expanding' | 'expanded' | 'reverting'>('idle')
  
  useEffect(() => {
    if (animationPhase === 'collapsing') {
      const timer = setTimeout(() => setAnimationPhase('expanding'), 600)
      return () => clearTimeout(timer)
    }
    if (animationPhase === 'expanding') {
      const timer = setTimeout(() => setAnimationPhase('expanded'), 800)
      return () => clearTimeout(timer)
    }
    if (animationPhase === 'reverting') {
      const timer = setTimeout(() => {
        setAnimationPhase('idle')
        setExpandedPlan(null)
        setIsTransitioning(false)
      }, 800)
      return () => clearTimeout(timer)
    }
    if (animationPhase === 'expanded') {
      setIsTransitioning(false)
    }
  }, [animationPhase])
  
  const handlePlanClick = (planId: string) => {
    if (isTransitioning) return
    
    if (expandedPlan === planId) {
      // Closing
      setIsTransitioning(true)
      setAnimationPhase('reverting')
    } else if (expandedPlan) {
      // Switching between plans
      setIsTransitioning(true)
      setAnimationPhase('reverting')
      setTimeout(() => {
        setExpandedPlan(planId)
        setAnimationPhase('expanding')
      }, 800)
    } else {
      // Opening
      setIsTransitioning(true)
      setExpandedPlan(planId)
      setAnimationPhase('collapsing')
    }
  }
  
  return (
    <section id="pricing" className="snap-section relative overflow-hidden bg-transparent">
      <style jsx>{`
        @keyframes slideAndScale {
          0% {
            transform: translateX(0) scale(1);
            opacity: 1;
          }
          50% {
            transform: translateX(-30px) scale(0.98);
            opacity: 1;
          }
          100% {
            transform: translateX(-40px) scale(1);
            opacity: 1;
          }
        }
        
        @keyframes slideFromFar {
          0% {
            transform: translateX(100px) scale(0.9);
            opacity: 0;
            filter: blur(4px);
          }
          50% {
            transform: translateX(50px) scale(0.95);
            opacity: 0.5;
            filter: blur(2px);
          }
          100% {
            transform: translateX(0) scale(1);
            opacity: 1;
            filter: blur(0);
          }
        }
        
        @keyframes gentleFadeOut {
          0% {
            opacity: 1;
            transform: scale(1) translateY(0);
          }
          100% {
            opacity: 0;
            transform: scale(0.92) translateY(10px);
            filter: blur(8px);
          }
        }
        
        @keyframes gentleFadeIn {
          0% {
            opacity: 0;
            transform: scale(1.08) translateY(-10px);
            filter: blur(8px);
          }
          100% {
            opacity: 1;
            transform: scale(1) translateY(0);
            filter: blur(0);
          }
        }
        
        @keyframes expandWidth {
          0% {
            max-width: 64rem;
          }
          100% {
            max-width: 80rem;
          }
        }
        
        @keyframes contractWidth {
          0% {
            max-width: 80rem;
          }
          100% {
            max-width: 64rem;
          }
        }
        
        @keyframes slideBack {
          0% {
            transform: translateX(-40px) scale(1);
          }
          50% {
            transform: translateX(-20px) scale(1.02);
          }
          100% {
            transform: translateX(0) scale(1);
          }
        }
        
        @keyframes slideAway {
          0% {
            transform: translateX(0) scale(1);
            opacity: 1;
            filter: blur(0);
          }
          100% {
            transform: translateX(100px) scale(0.9);
            opacity: 0;
            filter: blur(4px);
          }
        }
        
        .slide-and-scale {
          animation: slideAndScale 0.8s cubic-bezier(0.4, 0, 0.2, 1) forwards;
        }
        
        .slide-from-far {
          animation: slideFromFar 0.8s cubic-bezier(0.4, 0, 0.2, 1) forwards;
          animation-delay: 0.2s;
          opacity: 0;
        }
        
        .gentle-fade-out {
          animation: gentleFadeOut 0.6s cubic-bezier(0.4, 0, 0.2, 1) forwards;
        }
        
        .gentle-fade-in {
          animation: gentleFadeIn 0.6s cubic-bezier(0.4, 0, 0.2, 1) forwards;
        }
        
        .expand-container {
          animation: expandWidth 0.8s cubic-bezier(0.4, 0, 0.2, 1) forwards;
        }
        
        .contract-container {
          animation: contractWidth 0.8s cubic-bezier(0.4, 0, 0.2, 1) forwards;
        }
        
        .slide-back {
          animation: slideBack 0.8s cubic-bezier(0.4, 0, 0.2, 1) forwards;
        }
        
        .slide-away {
          animation: slideAway 0.8s cubic-bezier(0.4, 0, 0.2, 1) forwards;
        }
        
        .stagger-1 {
          animation-delay: 0.1s;
        }
        
        .stagger-2 {
          animation-delay: 0.2s;
        }
        
        .stagger-3 {
          animation-delay: 0.3s;
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

          {/* Expanded view */}
          {expandedPlan && (
            <div className={`mb-8 overflow-hidden transition-all duration-1000 ${
              animationPhase === 'expanding' || animationPhase === 'expanded' ? 'opacity-100' : 'opacity-0'
            }`}>
              {plans.filter(p => p.id === expandedPlan).map(plan => (
                <div key={plan.id} className={`mx-auto transition-all duration-800 ${
                  animationPhase === 'expanding' ? 'expand-container' : 
                  animationPhase === 'reverting' ? 'contract-container' : 
                  'max-w-5xl'
                }`}>
                  <div className="glass-card rounded-2xl overflow-hidden shadow-2xl">
                    <div className="grid lg:grid-cols-2 gap-0">
                      {/* Plan summary - slides left */}
                      <div className={`p-6 lg:p-8 ${
                        animationPhase === 'expanding' ? 'slide-and-scale' : 
                        animationPhase === 'reverting' ? 'slide-back' : ''
                      }`}>
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
                          className={`block w-full py-3 px-6 rounded-full text-center font-medium transition-all mb-4 transform hover:scale-105 ${
                            plan.popular
                              ? 'btn-primary'
                              : 'bg-white border border-neutral-300 text-neutral-700 hover:bg-neutral-100'
                          }`}
                        >
                          {plan.cta}
                        </Link>

                        <button
                          onClick={() => handlePlanClick(plan.id)}
                          className="text-sm text-neutral-500 hover:text-neutral-700 transition-all hover:translate-x-1"
                        >
                          ← Back to all plans
                        </button>
                      </div>

                      {/* Full feature list - slides in from right */}
                      <div className={`p-6 lg:p-8 bg-gradient-to-br from-neutral-50/50 to-white ${
                        animationPhase === 'expanding' ? 'slide-from-far' : 
                        animationPhase === 'reverting' ? 'slide-away' : ''
                      }`}>
                        <h4 className="text-lg font-semibold text-neutral-900 mb-4">All Features Included:</h4>
                        <ul className="space-y-3">
                          {plan.features.map((feature, index) => (
                            <li 
                              key={index} 
                              className={`flex items-start transform transition-all duration-500 ${
                                animationPhase === 'expanding' ? `stagger-${Math.min(index + 1, 3)}` : ''
                              }`}
                            >
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
          <div className={`grid grid-cols-1 lg:grid-cols-3 gap-6 lg:gap-8 items-stretch mb-6 sm:mb-8 transition-all duration-600 ${
            expandedPlan ? 
              (animationPhase === 'collapsing' ? 'gentle-fade-out' : 'opacity-0 pointer-events-none max-h-0 overflow-hidden') : 
              (animationPhase === 'reverting' ? 'gentle-fade-in' : 'opacity-100 max-h-[1000px]')
          }`}>
            {plans.map((plan, index) => (
              <ScrollObserver
                key={plan.name}
                className={`scroll-observer-delay-${index * 100}`}
              >
                <div
                  onClick={() => handlePlanClick(plan.id)}
                  className={`glass-card rounded-xl sm:rounded-2xl overflow-hidden h-full cursor-pointer transition-all duration-500 hover:scale-105 hover:shadow-2xl hover:-translate-y-2 ${
                    plan.popular ? 'ring-2 ring-primary-300' : ''
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
                      <span className="text-xs text-primary-600 font-medium hover:text-primary-700 transition-colors">
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