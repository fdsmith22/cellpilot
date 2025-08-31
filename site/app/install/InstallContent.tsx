'use client'

import { useState } from 'react'
import Link from 'next/link'
import ScrollObserver from '@/components/ScrollObserver'
import GridAnimation from '@/components/GridAnimation'
import Footer from '@/components/Footer'

type InstallMethod = 'marketplace' | 'manual' | null

export default function InstallContent() {
  const [selectedMethod, setSelectedMethod] = useState<InstallMethod>(null)
  
  return (
    <main className="min-h-screen relative">
      <GridAnimation />
      {/* Background decoration matching main site */}
      <div className="absolute inset-0 overflow-hidden pointer-events-none" style={{zIndex: 0}}>
        <div className="absolute top-1/4 right-0 w-[500px] h-[500px] bg-gradient-to-l from-pastel-sky/20 via-pastel-mint/10 to-transparent rounded-full blur-3xl"></div>
        <div className="absolute bottom-1/4 left-0 w-[500px] h-[500px] bg-gradient-to-r from-pastel-lavender/20 via-pastel-peach/10 to-transparent rounded-full blur-3xl"></div>
      </div>
      
      <div className="container mx-auto px-8 lg:px-6 md:px-4 max-md:px-4 pt-[104px] pb-20 max-sm:pb-16 relative" style={{zIndex: 10}}>
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
                  <p className="text-neutral-600">
                    The easiest way to install. One-click installation from the official marketplace.
                  </p>
                  <div className="mt-4 flex items-center text-sm text-primary-600">
                    <span>Recommended</span>
                    <svg className="w-4 h-4 ml-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                  </div>
                </div>
              </ScrollObserver>

              {/* Manual Method */}
              <ScrollObserver className="scroll-observer-delay-200">
                <div 
                  onClick={() => setSelectedMethod('manual')}
                  className="glass-card rounded-2xl p-6 sm:p-8 hover:scale-102 hover:shadow-xl transition-all cursor-pointer border-2 border-transparent hover:border-pastel-lavender/50 relative overflow-hidden group"
                >
                  {/* Cell reference */}
                  <div className="absolute -top-2 -left-2 text-[10px] text-neutral-500/70 font-mono">B1</div>
                  
                  <div className="flex items-center justify-center w-14 h-14 bg-gradient-to-br from-pastel-lavender/30 to-pastel-lavender/20 border border-pastel-lavender/40 rounded-xl mb-4 group-hover:scale-110 transition-transform">
                    <svg className="w-7 h-7 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4" />
                    </svg>
                  </div>
                  <h3 className="text-xl sm:text-2xl font-bold text-neutral-900 mb-3">
                    Manual Installation
                  </h3>
                  <p className="text-neutral-600">
                    Install directly from Google Sheets. Great for testing or restricted environments.
                  </p>
                  <div className="mt-4 flex items-center text-sm text-neutral-500">
                    <span>Advanced option</span>
                  </div>
                </div>
              </ScrollObserver>
            </div>
          ) : (
            <div className="glass-card rounded-2xl p-6 sm:p-8 mb-8">
              <button 
                onClick={() => setSelectedMethod(null)}
                className="mb-6 inline-flex items-center text-primary-600 hover:text-primary-700 transition-colors"
              >
                <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
                </svg>
                Back to methods
              </button>

              {selectedMethod === 'marketplace' ? (
                <div>
                  <h2 className="text-2xl sm:text-3xl font-bold text-neutral-900 mb-6">
                    Install from Marketplace
                  </h2>
                  
                  <div className="bg-amber-50 border border-amber-200 rounded-lg p-4 mb-6">
                    <p className="text-sm text-amber-800">
                      <strong>⏳ Coming Soon:</strong> CellPilot is currently under review for the Google Workspace Marketplace. For now, please use the manual installation method.
                    </p>
                  </div>
                  
                  <div className="space-y-8 opacity-50">
                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        1
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Visit the Marketplace</h3>
                        <p className="text-neutral-600 mb-3">Go to the Google Workspace Marketplace</p>
                        <button 
                          disabled
                          className="inline-flex items-center px-4 py-2 bg-neutral-300 text-neutral-500 rounded-lg cursor-not-allowed"
                        >
                          Coming Soon
                          <svg className="w-4 h-4 ml-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
                          </svg>
                        </button>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        2
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Search for CellPilot</h3>
                        <p className="text-neutral-600">Type "CellPilot" in the search bar and press Enter</p>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        3
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Click Install</h3>
                        <p className="text-neutral-600">Click the "Install" button and follow the prompts to grant permissions</p>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-green-100 rounded-full flex items-center justify-center text-green-600">
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                        </svg>
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Done!</h3>
                        <p className="text-neutral-600">CellPilot will now appear in your Google Sheets Extensions menu</p>
                      </div>
                    </div>
                  </div>
                  
                  <div className="mt-8 p-4 bg-blue-50 border border-blue-200 rounded-lg">
                    <p className="text-sm text-blue-800">
                      <strong>💡 Tip:</strong> Use the <button 
                        onClick={() => setSelectedMethod('manual')}
                        className="underline font-semibold hover:text-blue-900"
                      >manual installation method</button> to get started with CellPilot today!
                    </p>
                  </div>
                </div>
              ) : (
                <div>
                  <h2 className="text-2xl sm:text-3xl font-bold text-neutral-900 mb-6">
                    Manual Installation (Beta Access)
                  </h2>
                  
                  <div className="bg-blue-50 border border-blue-200 rounded-lg p-4 mb-6">
                    <p className="text-sm text-blue-800">
                      <strong>🚀 Beta Version:</strong> Install CellPilot as a library while we await marketplace approval
                    </p>
                  </div>
                  
                  <div className="space-y-8">
                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        1
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Open Google Sheets</h3>
                        <p className="text-neutral-600">Open any Google Sheets document or create a new one</p>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        2
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Open Apps Script Editor</h3>
                        <p className="text-neutral-600 mb-2">Click <strong>Extensions</strong> → <strong>Apps Script</strong></p>
                        <p className="text-sm text-neutral-500">This opens the script editor in a new tab</p>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        3
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Clear Default Code</h3>
                        <p className="text-neutral-600 mb-3">Delete any existing code in the editor (usually just a few lines)</p>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        4
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Add Bridge Code</h3>
                        <p className="text-neutral-600 mb-3">Copy and paste this code into the editor:</p>
                        <div className="bg-neutral-900 text-neutral-100 p-4 rounded-lg overflow-x-auto">
                          <pre className="text-xs sm:text-sm font-mono">{`function onOpen() {
  CellPilot.onOpen();
}

function onInstall(e) {
  CellPilot.onInstall(e);
}`}</pre>
                        </div>
                        <button 
                          onClick={() => {
                            navigator.clipboard.writeText(`function onOpen() {
  CellPilot.onOpen();
}

function onInstall(e) {
  CellPilot.onInstall(e);
}`);
                            // You could add a toast notification here
                          }}
                          className="mt-2 text-sm text-primary-600 hover:text-primary-700 flex items-center"
                        >
                          <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
                          </svg>
                          Copy code
                        </button>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        5
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Add CellPilot Library</h3>
                        <p className="text-neutral-600 mb-3">
                          Click the <strong>Libraries</strong> button (+ icon) in the left sidebar, then:
                        </p>
                        <ol className="list-decimal list-inside space-y-2 text-neutral-600">
                          <li>In the "Script ID" field, paste this ID:</li>
                        </ol>
                        <div className="bg-neutral-900 text-neutral-100 p-3 rounded-lg mt-2 mb-3">
                          <code className="text-xs sm:text-sm break-all">1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O</code>
                        </div>
                        <button 
                          onClick={() => {
                            navigator.clipboard.writeText('1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O');
                          }}
                          className="mb-3 text-sm text-primary-600 hover:text-primary-700 flex items-center"
                        >
                          <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
                          </svg>
                          Copy Script ID
                        </button>
                        <ol className="list-decimal list-inside space-y-2 text-neutral-600" start={2}>
                          <li>Click <strong>"Look up"</strong></li>
                          <li>Select the latest version</li>
                          <li>Make sure the identifier is set to <strong>"CellPilot"</strong></li>
                          <li>Click <strong>"Add"</strong></li>
                        </ol>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-primary-100 rounded-full flex items-center justify-center text-primary-600 font-semibold text-sm">
                        6
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Save and Authorize</h3>
                        <p className="text-neutral-600 mb-3">
                          Click the <strong>Save</strong> button (💾) and then <strong>Run</strong> → <strong>onInstall</strong>
                        </p>
                        <p className="text-neutral-600">
                          You'll be prompted to authorize. Click through the authorization screens.
                        </p>
                        <div className="mt-3 p-3 bg-amber-50 border border-amber-200 rounded-lg">
                          <p className="text-sm text-amber-800">
                            <strong>Note:</strong> You may see "This app isn't verified" - click "Advanced" → "Go to CellPilot (unsafe)" to proceed. This is normal for beta apps.
                          </p>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-start">
                      <div className="flex-shrink-0 w-8 h-8 bg-green-100 rounded-full flex items-center justify-center text-green-600">
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                        </svg>
                      </div>
                      <div className="ml-4">
                        <h3 className="font-semibold text-neutral-900 mb-2">Ready to Use!</h3>
                        <p className="text-neutral-600 mb-2">
                          Close the Apps Script tab and refresh your Google Sheet. 
                        </p>
                        <p className="text-neutral-600">
                          You'll now see <strong>CellPilot</strong> in your Extensions menu!
                        </p>
                      </div>
                    </div>
                  </div>
                  
                  <div className="mt-8 p-4 bg-green-50 border border-green-200 rounded-lg">
                    <h4 className="font-semibold text-green-900 mb-2">🎉 Beta Benefits</h4>
                    <ul className="space-y-1 text-sm text-green-800">
                      <li>• All features unlocked during beta period</li>
                      <li>• Help shape the product with your feedback</li>
                      <li>• Special pricing when we launch</li>
                    </ul>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Quick Start Guide */}
          <ScrollObserver>
            <div className="glass-card rounded-2xl p-6 sm:p-8">
              <h2 className="text-xl sm:text-2xl font-bold text-neutral-900 mb-4">
                After Installation
              </h2>
              <div className="space-y-4">
                <div className="flex items-start">
                  <svg className="w-5 h-5 text-green-500 mt-0.5 mr-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  <div>
                    <h3 className="font-semibold text-neutral-900">Open CellPilot Sidebar</h3>
                    <p className="text-sm text-neutral-600">Go to Extensions → CellPilot → Open Sidebar</p>
                  </div>
                </div>
                <div className="flex items-start">
                  <svg className="w-5 h-5 text-green-500 mt-0.5 mr-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  <div>
                    <h3 className="font-semibold text-neutral-900">Select Your Data</h3>
                    <p className="text-sm text-neutral-600">Highlight the cells you want to work with</p>
                  </div>
                </div>
                <div className="flex items-start">
                  <svg className="w-5 h-5 text-green-500 mt-0.5 mr-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  <div>
                    <h3 className="font-semibold text-neutral-900">Choose an Operation</h3>
                    <p className="text-sm text-neutral-600">Pick from our tools to clean, analyze, or transform your data</p>
                  </div>
                </div>
              </div>
            </div>
          </ScrollObserver>
        </div>
      </div>
      <Footer />
    </main>
  )
}