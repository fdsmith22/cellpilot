'use client'

import ScrollObserver from './ScrollObserver'

const features = [
  {
    title: 'Remove Duplicates',
    description: 'Find and remove duplicate entries in your spreadsheets with customizable matching options.',
    icon: (
      <svg className="h-6 w-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
      </svg>
    ),
  },
  {
    title: 'Formula Builder',
    description: 'Create complex formulas by describing what you need in plain language.',
    icon: (
      <svg className="h-6 w-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z" />
      </svg>
    ),
  },
  {
    title: 'Clean Text Data',
    description: 'Standardize text formatting, fix spacing issues, and clean up messy data.',
    icon: (
      <svg className="h-6 w-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
      </svg>
    ),
  },
  {
    title: 'Automatic Backups',
    description: 'Your data is automatically backed up before any operation so you can undo changes.',
    icon: (
      <svg className="h-6 w-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
      </svg>
    ),
  },
  {
    title: 'Email Alerts',
    description: 'Set up automated email notifications based on your spreadsheet data.',
    icon: (
      <svg className="h-6 w-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
      </svg>
    ),
  },
  {
    title: 'Usage Tracking',
    description: 'Monitor your usage and see how many operations you have remaining.',
    icon: (
      <svg className="h-6 w-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
      </svg>
    ),
  },
]

const Features = () => {
  return (
    <section id="features" className="snap-section h-screen relative overflow-hidden bg-transparent">
      {/* Background decoration */}
      <div className="absolute inset-0">
        <div className="absolute top-0 right-0 w-1/3 h-1/3 bg-pastel-sky/10 rounded-full blur-3xl"></div>
        <div className="absolute bottom-0 left-0 w-1/3 h-1/3 bg-pastel-mint/10 rounded-full blur-3xl"></div>
      </div>
      
      {/* Seamless transition from hero */}
      <div className="absolute inset-0 bg-gradient-to-b from-transparent via-white/50 to-white/60"></div>
      
      {/* Removed grid pattern - using global grid */}
      
      <div className="container-wrapper relative z-10 flex items-center justify-center h-screen pt-[104px] pb-20">
        <div className="w-full">
          <ScrollObserver className="text-center max-w-3xl mx-auto mb-8 sm:mb-12 px-4">
          <h2 className="text-2xl sm:text-3xl md:text-4xl lg:text-5xl font-bold text-neutral-900 mb-3 sm:mb-4">
            Tools That Actually Help
          </h2>
          <p className="text-sm sm:text-base md:text-lg text-neutral-600 leading-relaxed">
            Simple, practical features to make working with Google Sheets easier. 
            No complicated setup, just tools that work.
          </p>
        </ScrollObserver>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 sm:gap-6 px-4 sm:px-0">
          {features.map((feature, index) => (
            <ScrollObserver
              key={index}
              className={`scroll-observer-delay-${index * 100}`}
            >
              <div className="group glass-card rounded-xl sm:rounded-2xl p-6 sm:p-8 card-hover h-full">
              <div className="inline-flex p-2 sm:p-3 rounded-lg sm:rounded-xl bg-gradient-to-br from-pastel-lavender/20 to-pastel-sky/20 text-primary-600 mb-3 sm:mb-4">
                {feature.icon}
              </div>
              
              <h3 className="text-lg sm:text-xl font-semibold text-neutral-900 mb-2 sm:mb-3">
                {feature.title}
              </h3>
              
              <p className="text-sm sm:text-base text-neutral-600 leading-relaxed">
                {feature.description}
              </p>
              </div>
            </ScrollObserver>
          ))}
          </div>
        </div>
      </div>
    </section>
  )
}

export default Features