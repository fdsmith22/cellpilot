import Link from 'next/link'
import ScrollObserver from './ScrollObserver'

const CTA = () => {
  return (
    <section className="snap-section h-screen-vh md:h-screen relative overflow-hidden bg-transparent">
      {/* Animated background elements */}
      <div className="absolute inset-0 overflow-hidden">
        <div className="absolute top-20 left-20 w-64 h-64 bg-pastel-sky/20 rounded-full blur-2xl animate-blob"></div>
        <div className="absolute bottom-20 right-20 w-64 h-64 bg-pastel-mint/20 rounded-full blur-2xl animate-blob animation-delay-4000"></div>
      </div>
      
      {/* Seamless transition */}
      <div className="absolute inset-0 bg-gradient-to-b from-white/70 via-primary-50/20 to-white/90"></div>
      
      {/* Removed grid pattern - using global grid */}
      
      <div className="container-wrapper relative z-elevated h-screen-vh md:h-screen pt-header-total-mobile md:pt-header-total pb-safe-bottom md:pb-4 flex items-center justify-center">
        <ScrollObserver className="max-w-3xl mx-auto text-center px-4">
          <h2 className="text-2xl sm:text-3xl md:text-4xl lg:text-5xl font-bold text-neutral-900 mb-4 sm:mb-6">
            Ready to Get Started?
          </h2>
          
          <p className="text-base sm:text-lg md:text-xl text-neutral-600 mb-6 sm:mb-10 leading-relaxed">
            Try CellPilot free and see how it can help streamline your Google Sheets workflow.
          </p>
          
          <div className="flex flex-col sm:flex-row gap-4 justify-center">
            <Link
              href="/install"
              className="btn-primary"
            >
              <span>Install Free Add-on</span>
              <svg className="w-4 h-4 ml-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
              </svg>
            </Link>
            <Link
              href="/docs"
              className="btn-secondary"
            >
              <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253" />
              </svg>
              <span>Read Documentation</span>
            </Link>
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