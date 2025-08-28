'use client'

import { useState } from 'react'
import Link from 'next/link'
import ScrollObserver from '@/components/ScrollObserver'
import GridAnimation from '@/components/GridAnimation'

type InstallMethod = 'marketplace' | 'manual' | null

export default function InstallPage() {
  const [selectedMethod, setSelectedMethod] = useState<InstallMethod>(null)
  
  return (
    <main className="min-h-screen relative">
      <GridAnimation />
      {/* Background decoration matching main site */}
      <div className="absolute inset-0 overflow-hidden pointer-events-none" style={{zIndex: 0}}>
        <div className="absolute top-1/4 right-0 w-[500px] h-[500px] bg-gradient-to-l from-pastel-sky/20 via-pastel-mint/10 to-transparent rounded-full blur-3xl"></div>
        <div className="absolute bottom-1/4 left-0 w-[500px] h-[500px] bg-gradient-to-r from-pastel-lavender/20 via-pastel-peach/10 to-transparent rounded-full blur-3xl"></div>
      </div>
      
      <div className="container mx-auto px-8 lg:px-6 md:px-4 max-md:px-4 pt-[100px] sm:pt-[110px] md:pt-[120px] lg:pt-[130px] pb-20 max-sm:pb-16 relative" style={{zIndex: 10}}>
        <div className="max-w-4xl mx-auto">
          <ScrollObserver className="text-center mb-12">
            <h1 className="text-4xl sm:text-5xl md:text-6xl font-bold text-neutral-900 mb-4">
              Install CellPilot
            </h1>
            <p className="text-lg sm:text-xl text-neutral-600">
              Choose your preferred installation method
            </p>
          </ScrollObserver>

          {!selectedMethod ? (
            <div className="grid md:grid-cols-2 gap-6 sm:gap-8 mb-12">
              {/* Marketplace Method */}
              <ScrollObserver className="scroll-observer-delay-100">
                <div 
                  onClick={() => setSelectedMethod('marketplace')}
                  className="glass-card rounded-2xl p-6 sm:p-8 hover:scale-102 hover:shadow-xl transition-all cursor-pointer border-2 border-transparent hover:border-pastel-sky/50 relative overflow-hidden group"
                >
                  {/* Cell reference */}
                  <div className="absolute -top-2 -left-2 text-[10px] text-neutral-500/70 font-mono">A1</div>
                  
                  <div className="flex items-center justify-center w-14 h-14 bg-gradient-to-br from-pastel-sky/30 to-pastel-sky/20 border border-pastel-sky/40 rounded-xl mb-4 group-hover:scale-110 transition-transform">
                    <svg className="w-7 h-7 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 3h2l.4 2M7 13h10l4-8H5.4M7 13L5.4 5M7 13l-2.293 2.293c-.63.63-.184 1.707.707 1.707H17m0 0a2 2 0 100 4 2 2 0 000-4zm-8 2a2 2 0 11-4 0 2 2 0 014 0z" />
                    </svg>
                  </div>
                  <h3 className="text-xl sm:text-2xl font-bold text-neutral-900 mb-3">
                    Google Workspace Marketplace
                  </h3>
                  <p className="text-neutral-600 mb-4">
                    Recommended for most users. Simple one-click installation.
                  </p>
                  <ul className="space-y-2.5">
                    <li className="flex items-start">
                      <svg className="w-5 h-5 text-accent-teal mr-2.5 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                      </svg>
                      <span className="text-sm text-neutral-700">One-click installation</span>
                    </li>
                    <li className="flex items-start">
                      <svg className="w-5 h-5 text-accent-teal mr-2.5 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                      </svg>
                      <span className="text-sm text-neutral-700">Automatic updates</span>
                    </li>
                    <li className="flex items-start">
                      <svg className="w-5 h-5 text-accent-teal mr-2.5 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                      </svg>
                      <span className="text-sm text-neutral-700">Google verified</span>
                    </li>
                  </ul>
                </div>
              </ScrollObserver>

              {/* Manual Method */}
              <ScrollObserver className="scroll-observer-delay-200">
                <div 
                  onClick={() => setSelectedMethod('manual')}
                  className="glass-card rounded-2xl p-6 sm:p-8 hover:scale-102 hover:shadow-xl transition-all cursor-pointer border-2 border-transparent hover:border-pastel-mint/50 relative overflow-hidden group"
                >
                  {/* Cell reference */}
                  <div className="absolute -top-2 -left-2 text-[10px] text-neutral-500/70 font-mono">B1</div>
                  
                  <div className="flex items-center justify-center w-14 h-14 bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 border border-pastel-mint/40 rounded-xl mb-4 group-hover:scale-110 transition-transform">
                    <svg className="w-7 h-7 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 20l4-16m4 4l4 4-4 4M6 16l-4-4 4-4" />
                    </svg>
                  </div>
                  <h3 className="text-xl sm:text-2xl font-bold text-neutral-900 mb-3">
                    Manual Installation
                  </h3>
                  <p className="text-neutral-600 mb-4">
                    For developers and power users. Full control over the code.
                  </p>
                  <ul className="space-y-2.5">
                    <li className="flex items-start">
                      <svg className="w-5 h-5 text-accent-teal mr-2.5 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                      </svg>
                      <span className="text-sm text-neutral-700">Full code access</span>
                    </li>
                    <li className="flex items-start">
                      <svg className="w-5 h-5 text-accent-teal mr-2.5 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                      </svg>
                      <span className="text-sm text-neutral-700">Customizable</span>
                    </li>
                    <li className="flex items-start">
                      <svg className="w-5 h-5 text-accent-teal mr-2.5 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                      </svg>
                      <span className="text-sm text-neutral-700">No permissions needed</span>
                    </li>
                  </ul>
                </div>
              </ScrollObserver>
            </div>
          ) : (
            <div className="glass-card rounded-2xl p-6 sm:p-8">
              <button 
                onClick={() => setSelectedMethod(null)}
                className="mb-6 text-sm text-neutral-600 hover:text-primary-600 flex items-center transition-colors"
              >
                <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
                </svg>
                Back to methods
              </button>

              {selectedMethod === 'marketplace' ? (
                <MarketplaceInstructions />
              ) : (
                <ManualInstructions />
              )}
            </div>
          )}
        </div>
      </div>
    </main>
  )
}

function MarketplaceInstructions() {
  return (
    <>
      <h2 className="text-2xl sm:text-3xl font-bold text-neutral-900 mb-6">
        Install from Google Workspace Marketplace
      </h2>
      
      <div className="space-y-6">
        <div className="bg-gradient-to-r from-pastel-sky/20 to-pastel-sky/10 border border-pastel-sky/40 rounded-xl p-4">
          <p className="text-neutral-700">
            <strong>Note:</strong> CellPilot is currently in review for the Google Workspace Marketplace. 
            Expected availability: January 2025. Use manual installation for immediate access.
          </p>
        </div>

        <ol className="space-y-4">
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-sky/30 to-pastel-sky/20 border border-pastel-sky/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">1</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Open Google Workspace Marketplace</h3>
              <p className="text-neutral-600 mt-1">In any Google Sheet, go to Extensions → Add-ons → Get add-ons</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-sky/30 to-pastel-sky/20 border border-pastel-sky/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">2</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Search for CellPilot</h3>
              <p className="text-neutral-600 mt-1">Type "CellPilot" in the search bar and select our add-on</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-sky/30 to-pastel-sky/20 border border-pastel-sky/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">3</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Click Install</h3>
              <p className="text-neutral-600 mt-1">Review permissions and click "Install" to add CellPilot to your account</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-sky/30 to-pastel-sky/20 border border-pastel-sky/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">4</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Start using CellPilot</h3>
              <p className="text-neutral-600 mt-1">Refresh your Google Sheet and find CellPilot in the Extensions menu</p>
            </div>
          </li>
        </ol>

        <div className="glass-card rounded-lg p-6 mt-8">
          <h3 className="font-semibold text-neutral-900 mb-2">Permissions Required</h3>
          <ul className="space-y-2 text-sm text-neutral-600">
            <li className="flex items-start">
              <svg className="w-5 h-5 text-accent-teal mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              View and manage spreadsheets that this application has been installed in
            </li>
            <li className="flex items-start">
              <svg className="w-5 h-5 text-accent-teal mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              Display and run third-party web content in prompts and sidebars
            </li>
          </ul>
        </div>
      </div>
    </>
  )
}

function ManualInstructions() {
  const scriptId = "1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O"
  
  return (
    <>
      <h2 className="text-2xl sm:text-3xl font-bold text-neutral-900 mb-6">
        Manual Installation Instructions
      </h2>
      
      <div className="space-y-6">
        <div className="bg-gradient-to-r from-pastel-mint/20 to-pastel-mint/10 border border-pastel-mint/40 rounded-xl p-4">
          <p className="text-neutral-700">
            <strong>Quick Install:</strong> Copy the script ID below and add CellPilot as a library to your Google Sheets.
          </p>
        </div>

        <div className="glass-card rounded-lg p-4">
          <label className="block text-sm font-medium text-neutral-700 mb-2">Script ID</label>
          <div className="flex items-center space-x-2">
            <input
              type="text"
              value={scriptId}
              readOnly
              className="flex-1 px-3 py-2 border border-neutral-300 rounded-lg bg-white text-sm font-mono"
            />
            <button
              onClick={() => navigator.clipboard.writeText(scriptId)}
              className="px-4 py-2 bg-gradient-to-br from-primary-500 to-primary-600 text-white rounded-lg hover:from-primary-600 hover:to-primary-700 hover:shadow-lg transition-all"
            >
              Copy
            </button>
          </div>
        </div>

        <ol className="space-y-4">
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 border border-pastel-mint/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">1</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Open Apps Script Editor</h3>
              <p className="text-neutral-600 mt-1">In your Google Sheet, go to Extensions → Apps Script</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 border border-pastel-mint/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">2</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Add CellPilot as Library</h3>
              <p className="text-neutral-600 mt-1">Click the "+" next to Libraries, paste the Script ID, and click "Look up"</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 border border-pastel-mint/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">3</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Select Version and Add</h3>
              <p className="text-neutral-600 mt-1">Choose the latest version, keep identifier as "CellPilot", and click "Add"</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 border border-pastel-mint/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">4</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Initialize CellPilot</h3>
              <p className="text-neutral-600 mt-1">Replace the default code with:</p>
              <pre className="mt-2 p-3 bg-neutral-900 text-neutral-100 rounded-lg text-sm overflow-x-auto border border-neutral-800">
{`function onOpen() {
  CellPilot.onOpen();
}

function showCellPilotSidebar() {
  CellPilot.showCellPilotSidebar();
}`}
              </pre>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gradient-to-br from-pastel-mint/30 to-pastel-mint/20 border border-pastel-mint/40 text-primary-700 rounded-full flex items-center justify-center font-semibold">5</span>
            <div className="ml-4">
              <h3 className="font-semibold text-neutral-900">Save and Refresh</h3>
              <p className="text-neutral-600 mt-1">Save your script (Ctrl+S) and refresh your Google Sheet to see the CellPilot menu</p>
            </div>
          </li>
        </ol>

        <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4 mt-6">
          <h3 className="font-semibold text-yellow-900 mb-2">Advanced Users</h3>
          <p className="text-yellow-800 text-sm">
            Want to customize CellPilot? You can clone the entire repository and modify it to your needs.
          </p>
          <a 
            href="https://github.com/cellpilot/cellpilot-sheets"
            className="inline-block mt-2 text-yellow-900 hover:text-yellow-700 font-medium text-sm"
            target="_blank"
            rel="noopener noreferrer"
          >
            View on GitHub →
          </a>
        </div>
      </div>
    </>
  )
}