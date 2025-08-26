'use client'

import { useState } from 'react'
import Link from 'next/link'

type InstallMethod = 'marketplace' | 'manual' | null

export default function InstallPage() {
  const [selectedMethod, setSelectedMethod] = useState<InstallMethod>(null)
  
  return (
    <main className="min-h-screen bg-gradient-to-b from-gray-50 to-white">
      <div className="container mx-auto px-4 py-16 sm:px-6 lg:px-8">
        <div className="max-w-4xl mx-auto">
          <div className="text-center mb-12">
            <h1 className="text-4xl font-bold text-gray-900 mb-4">
              Install CellPilot
            </h1>
            <p className="text-lg text-gray-600">
              Choose your preferred installation method
            </p>
          </div>

          {!selectedMethod ? (
            <div className="grid md:grid-cols-2 gap-8 mb-12">
              {/* Marketplace Method */}
              <div 
                onClick={() => setSelectedMethod('marketplace')}
                className="bg-white rounded-lg border-2 border-gray-200 p-8 hover:border-primary-500 hover:shadow-lg transition-all cursor-pointer"
              >
                <div className="flex items-center justify-center w-12 h-12 bg-primary-100 text-primary-600 rounded-lg mb-4">
                  <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 3h2l.4 2M7 13h10l4-8H5.4M7 13L5.4 5M7 13l-2.293 2.293c-.63.63-.184 1.707.707 1.707H17m0 0a2 2 0 100 4 2 2 0 000-4zm-8 2a2 2 0 11-4 0 2 2 0 014 0z" />
                  </svg>
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">
                  Google Workspace Marketplace
                </h3>
                <p className="text-gray-600 mb-4">
                  Recommended for most users. Simple one-click installation.
                </p>
                <ul className="space-y-2 text-sm text-gray-500">
                  <li className="flex items-start">
                    <svg className="w-5 h-5 text-green-500 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    One-click installation
                  </li>
                  <li className="flex items-start">
                    <svg className="w-5 h-5 text-green-500 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    Automatic updates
                  </li>
                  <li className="flex items-start">
                    <svg className="w-5 h-5 text-green-500 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    Google verified
                  </li>
                </ul>
              </div>

              {/* Manual Method */}
              <div 
                onClick={() => setSelectedMethod('manual')}
                className="bg-white rounded-lg border-2 border-gray-200 p-8 hover:border-primary-500 hover:shadow-lg transition-all cursor-pointer"
              >
                <div className="flex items-center justify-center w-12 h-12 bg-gray-100 text-gray-600 rounded-lg mb-4">
                  <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 20l4-16m4 4l4 4-4 4M6 16l-4-4 4-4" />
                  </svg>
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">
                  Manual Installation
                </h3>
                <p className="text-gray-600 mb-4">
                  For developers and power users. Full control over the code.
                </p>
                <ul className="space-y-2 text-sm text-gray-500">
                  <li className="flex items-start">
                    <svg className="w-5 h-5 text-green-500 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    Full code access
                  </li>
                  <li className="flex items-start">
                    <svg className="w-5 h-5 text-green-500 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    Customizable
                  </li>
                  <li className="flex items-start">
                    <svg className="w-5 h-5 text-green-500 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    No permissions needed
                  </li>
                </ul>
              </div>
            </div>
          ) : (
            <div className="bg-white rounded-lg shadow-lg p-8">
              <button 
                onClick={() => setSelectedMethod(null)}
                className="mb-6 text-sm text-gray-600 hover:text-primary-600 flex items-center"
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
      <h2 className="text-2xl font-bold text-gray-900 mb-6">
        Install from Google Workspace Marketplace
      </h2>
      
      <div className="space-y-6">
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
          <p className="text-blue-800">
            <strong>Note:</strong> CellPilot is currently in review for the Google Workspace Marketplace. 
            Expected availability: January 2025. Use manual installation for immediate access.
          </p>
        </div>

        <ol className="space-y-4">
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-primary-100 text-primary-600 rounded-full flex items-center justify-center font-semibold">1</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Open Google Workspace Marketplace</h3>
              <p className="text-gray-600 mt-1">In any Google Sheet, go to Extensions → Add-ons → Get add-ons</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-primary-100 text-primary-600 rounded-full flex items-center justify-center font-semibold">2</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Search for CellPilot</h3>
              <p className="text-gray-600 mt-1">Type "CellPilot" in the search bar and select our add-on</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-primary-100 text-primary-600 rounded-full flex items-center justify-center font-semibold">3</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Click Install</h3>
              <p className="text-gray-600 mt-1">Review permissions and click "Install" to add CellPilot to your account</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-primary-100 text-primary-600 rounded-full flex items-center justify-center font-semibold">4</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Start using CellPilot</h3>
              <p className="text-gray-600 mt-1">Refresh your Google Sheet and find CellPilot in the Extensions menu</p>
            </div>
          </li>
        </ol>

        <div className="bg-gray-50 rounded-lg p-6 mt-8">
          <h3 className="font-semibold text-gray-900 mb-2">Permissions Required</h3>
          <ul className="space-y-2 text-sm text-gray-600">
            <li className="flex items-start">
              <svg className="w-5 h-5 text-gray-400 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
              View and manage spreadsheets that this application has been installed in
            </li>
            <li className="flex items-start">
              <svg className="w-5 h-5 text-gray-400 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
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
      <h2 className="text-2xl font-bold text-gray-900 mb-6">
        Manual Installation Instructions
      </h2>
      
      <div className="space-y-6">
        <div className="bg-green-50 border border-green-200 rounded-lg p-4">
          <p className="text-green-800">
            <strong>Quick Install:</strong> Copy the script ID below and add CellPilot as a library to your Google Sheets.
          </p>
        </div>

        <div className="bg-gray-50 rounded-lg p-4">
          <label className="block text-sm font-medium text-gray-700 mb-2">Script ID</label>
          <div className="flex items-center space-x-2">
            <input
              type="text"
              value={scriptId}
              readOnly
              className="flex-1 px-3 py-2 border border-gray-300 rounded-lg bg-white text-sm font-mono"
            />
            <button
              onClick={() => navigator.clipboard.writeText(scriptId)}
              className="px-4 py-2 bg-primary-600 text-white rounded-lg hover:bg-primary-700 transition-colors"
            >
              Copy
            </button>
          </div>
        </div>

        <ol className="space-y-4">
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gray-100 text-gray-600 rounded-full flex items-center justify-center font-semibold">1</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Open Apps Script Editor</h3>
              <p className="text-gray-600 mt-1">In your Google Sheet, go to Extensions → Apps Script</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gray-100 text-gray-600 rounded-full flex items-center justify-center font-semibold">2</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Add CellPilot as Library</h3>
              <p className="text-gray-600 mt-1">Click the "+" next to Libraries, paste the Script ID, and click "Look up"</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gray-100 text-gray-600 rounded-full flex items-center justify-center font-semibold">3</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Select Version and Add</h3>
              <p className="text-gray-600 mt-1">Choose the latest version, keep identifier as "CellPilot", and click "Add"</p>
            </div>
          </li>
          
          <li className="flex">
            <span className="flex-shrink-0 w-8 h-8 bg-gray-100 text-gray-600 rounded-full flex items-center justify-center font-semibold">4</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Initialize CellPilot</h3>
              <p className="text-gray-600 mt-1">Replace the default code with:</p>
              <pre className="mt-2 p-3 bg-gray-900 text-gray-100 rounded text-sm overflow-x-auto">
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
            <span className="flex-shrink-0 w-8 h-8 bg-gray-100 text-gray-600 rounded-full flex items-center justify-center font-semibold">5</span>
            <div className="ml-4">
              <h3 className="font-semibold text-gray-900">Save and Refresh</h3>
              <p className="text-gray-600 mt-1">Save your script (Ctrl+S) and refresh your Google Sheet to see the CellPilot menu</p>
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