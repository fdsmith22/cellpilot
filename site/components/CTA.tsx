import Link from 'next/link'

const CTA = () => {
  return (
    <section className="py-20 bg-primary-600">
      <div className="container mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h2 className="text-3xl font-bold text-white sm:text-4xl">
          Ready to transform your Google Sheets workflow?
        </h2>
        <p className="mt-4 text-xl text-primary-100">
          Join thousands of users saving hours every week with CellPilot
        </p>
        
        <div className="mt-10 flex flex-col sm:flex-row gap-4 justify-center">
          <Link
            href="/install"
            className="inline-block rounded-lg bg-white px-8 py-3 text-primary-600 font-semibold hover:bg-gray-100 transition-colors"
          >
            Get Started Free
          </Link>
          <Link
            href="/docs"
            className="inline-block rounded-lg border-2 border-white px-8 py-3 text-white font-semibold hover:bg-primary-700 transition-colors"
          >
            View Documentation
          </Link>
        </div>

        <div className="mt-8 flex items-center justify-center gap-6 text-white/90">
          <div className="flex items-center gap-2">
            <svg className="h-5 w-5" fill="currentColor" viewBox="0 0 20 20">
              <path d="M9.049 2.927c.3-.921 1.603-.921 1.902 0l1.07 3.292a1 1 0 00.95.69h3.462c.969 0 1.371 1.24.588 1.81l-2.8 2.034a1 1 0 00-.364 1.118l1.07 3.292c.3.921-.755 1.688-1.54 1.118l-2.8-2.034a1 1 0 00-1.175 0l-2.8 2.034c-.784.57-1.838-.197-1.539-1.118l1.07-3.292a1 1 0 00-.364-1.118L2.98 8.72c-.783-.57-.38-1.81.588-1.81h3.461a1 1 0 00.951-.69l1.07-3.292z" />
            </svg>
            <span>4.8/5 rating</span>
          </div>
          <div>|</div>
          <div>10,000+ active users</div>
          <div>|</div>
          <div>2M+ operations processed</div>
        </div>
      </div>
    </section>
  )
}

export default CTA