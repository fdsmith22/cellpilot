'use client'

import { useState } from 'react'
import Link from 'next/link'

const Hero = () => {
  const [email, setEmail] = useState('')

  return (
    <section className="relative overflow-hidden bg-gradient-to-b from-primary-50 to-white">
      <div className="container mx-auto px-4 py-24 sm:px-6 lg:px-8">
        <div className="text-center">
          <h1 className="text-4xl font-bold tracking-tight text-gray-900 sm:text-5xl md:text-6xl">
            Google Sheets Automation
            <span className="block text-primary-600">Made Simple</span>
          </h1>
          
          <p className="mx-auto mt-6 max-w-2xl text-lg text-gray-600">
            Transform your spreadsheet workflow with powerful data cleaning, 
            formula generation, and automation tools. Save hours of manual work 
            with CellPilot.
          </p>

          <div className="mt-10 flex flex-col sm:flex-row gap-4 justify-center">
            <Link
              href="/install"
              className="rounded-lg bg-primary-600 px-8 py-3 text-white font-semibold hover:bg-primary-700 transition-colors"
            >
              Get Started Free
            </Link>
            <Link
              href="#demo"
              className="rounded-lg border border-gray-300 bg-white px-8 py-3 text-gray-700 font-semibold hover:bg-gray-50 transition-colors"
            >
              Watch Demo
            </Link>
          </div>

          <div className="mt-8 flex items-center justify-center gap-8 text-sm text-gray-500">
            <div className="flex items-center gap-2">
              <svg className="h-5 w-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              No credit card required
            </div>
            <div className="flex items-center gap-2">
              <svg className="h-5 w-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              25 free operations/month
            </div>
            <div className="flex items-center gap-2">
              <svg className="h-5 w-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              5-minute setup
            </div>
          </div>
        </div>
      </div>
    </section>
  )
}

export default Hero