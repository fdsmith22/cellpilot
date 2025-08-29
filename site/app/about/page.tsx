import GridAnimation from '@/components/GridAnimation'
import ScrollObserver from '@/components/ScrollObserver'
import Link from 'next/link'

export default function AboutPage() {
  return (
    <>
      <GridAnimation />
      
      <div className="min-h-screen bg-gradient-to-b from-white/60 via-white/70 to-white/80">
        <div className="container-wrapper py-20 sm:py-24">
          <ScrollObserver>
            <div className="max-w-4xl mx-auto">
              <h1 className="text-4xl sm:text-5xl font-bold text-neutral-900 mb-8 text-center">
                About CellPilot
              </h1>
              
              <div className="prose prose-lg max-w-none space-y-8">
                <div className="glass-card rounded-2xl p-8">
                  <h2 className="text-2xl font-semibold text-neutral-900 mb-4">
                    Why We Built CellPilot
                  </h2>
                  <p className="text-neutral-700 leading-relaxed mb-4">
                    If you've ever spent hours trying to clean up a messy data export, struggled to write a formula that actually works, or found yourself manually copying and pasting the same data over and over, you're not alone. Google Sheets is powerful, but for most people, that power stays locked away behind complex formulas and hidden features.
                  </p>
                  <p className="text-neutral-700 leading-relaxed">
                    We built CellPilot because spreadsheet work shouldn't require a computer science degree. Whether you're dealing with unformatted data dumps that need splitting into proper rows and columns, trying to remove duplicates from a customer list, or simply want to clean up text formatting across thousands of cells, these are everyday tasks that should be simple.
                  </p>
                </div>

                <div className="glass-card rounded-2xl p-8">
                  <h2 className="text-2xl font-semibold text-neutral-900 mb-4">
                    What CellPilot Does
                  </h2>
                  <p className="text-neutral-700 leading-relaxed mb-6">
                    CellPilot sits right inside your Google Sheets, giving you practical tools that handle the tedious stuff so you can focus on what matters. Here's how we help:
                  </p>
                  
                  <div className="space-y-4">
                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-2 h-2 rounded-full bg-primary-500 mt-2.5 mr-3"></div>
                      <div>
                        <h3 className="font-semibold text-neutral-900 mb-1">Data Cleaning Made Simple</h3>
                        <p className="text-neutral-600">Transform messy paste dumps into organized data. Split combined fields, standardize formats, fix inconsistent capitalization, and clean up extra spaces with a single click.</p>
                      </div>
                    </div>
                    
                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-2 h-2 rounded-full bg-primary-500 mt-2.5 mr-3"></div>
                      <div>
                        <h3 className="font-semibold text-neutral-900 mb-1">Formula Builder</h3>
                        <p className="text-neutral-600">Describe what you want in plain English, and we'll generate the formula. No more googling syntax or debugging parentheses.</p>
                      </div>
                    </div>
                    
                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-2 h-2 rounded-full bg-primary-500 mt-2.5 mr-3"></div>
                      <div>
                        <h3 className="font-semibold text-neutral-900 mb-1">Duplicate Detection</h3>
                        <p className="text-neutral-600">Find and remove duplicate entries intelligently. Whether it's exact matches or similar entries with slight variations, we'll help you clean up your data.</p>
                      </div>
                    </div>
                    
                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-2 h-2 rounded-full bg-primary-500 mt-2.5 mr-3"></div>
                      <div>
                        <h3 className="font-semibold text-neutral-900 mb-1">Smart Text Processing</h3>
                        <p className="text-neutral-600">Extract emails from text blocks, split full names into first and last, standardize phone numbers, or parse addresses into separate fields.</p>
                      </div>
                    </div>
                    
                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-2 h-2 rounded-full bg-primary-500 mt-2.5 mr-3"></div>
                      <div>
                        <h3 className="font-semibold text-neutral-900 mb-1">Automation Templates</h3>
                        <p className="text-neutral-600">Set up recurring tasks once and run them whenever needed. Perfect for weekly reports, monthly data cleanups, or any repetitive process.</p>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="glass-card rounded-2xl p-8">
                  <h2 className="text-2xl font-semibold text-neutral-900 mb-4">
                    Who Uses CellPilot
                  </h2>
                  <p className="text-neutral-700 leading-relaxed mb-4">
                    Our users aren't spreadsheet experts, and that's exactly the point. They're sales teams organizing leads, HR departments managing employee data, small business owners tracking inventory, teachers grading assignments, and anyone else who uses Google Sheets as a tool, not a hobby.
                  </p>
                  <p className="text-neutral-700 leading-relaxed">
                    If spreadsheets are part of your work but not your passion, CellPilot is for you. We handle the technical complexity so you can get your work done and move on with your day.
                  </p>
                </div>

                <div className="glass-card rounded-2xl p-8">
                  <h2 className="text-2xl font-semibold text-neutral-900 mb-4">
                    Our Approach
                  </h2>
                  <p className="text-neutral-700 leading-relaxed mb-4">
                    We believe good tools should be invisible. You shouldn't need training videos or documentation to use CellPilot. Every feature is designed to be discoverable and intuitive. Click it, try it, and if it helps, great. If not, try something else.
                  </p>
                  <p className="text-neutral-700 leading-relaxed">
                    We're constantly adding new capabilities based on what our users actually need. Not flashy features for marketing pages, but practical tools that solve real problems you face every day.
                  </p>
                </div>

                <div className="text-center mt-12">
                  <Link
                    href="/install"
                    className="inline-flex items-center px-8 py-4 text-lg font-semibold text-white bg-gradient-to-r from-primary-600 to-primary-700 rounded-xl hover:from-primary-700 hover:to-primary-800 transition-all duration-300 shadow-lg hover:shadow-xl hover:scale-105"
                  >
                    Try CellPilot Free
                    <svg className="w-5 h-5 ml-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 8l4 4m0 0l-4 4m4-4H3" />
                    </svg>
                  </Link>
                  <p className="mt-4 text-sm text-neutral-600">
                    Start with 25 free operations per month. No credit card needed.
                  </p>
                </div>
              </div>
            </div>
          </ScrollObserver>
        </div>
      </div>
    </>
  )
}