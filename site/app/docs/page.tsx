import Link from 'next/link'
import type { Metadata } from 'next'
import GridAnimation from '@/components/GridAnimation'

export const metadata: Metadata = {
  title: 'Documentation - CellPilot',
  description: 'Complete documentation for using CellPilot Google Sheets automation add-on.',
}

export default function Documentation() {
  return (
    <main className="min-h-screen relative">
      <GridAnimation />
      {/* Background decoration matching main site */}
      <div className="absolute inset-0 overflow-hidden pointer-events-none" style={{zIndex: 0}}>
        <div className="absolute top-1/3 left-0 w-96 h-96 bg-pastel-mint/10 rounded-full blur-3xl"></div>
        <div className="absolute bottom-1/3 right-0 w-96 h-96 bg-pastel-lavender/10 rounded-full blur-3xl"></div>
      </div>
      
      <div className="container-wrapper relative pt-[104px] pb-20" style={{zIndex: 10}}>
        <div className="max-w-5xl mx-auto">
          {/* Back link */}
          <Link href="/" className="inline-flex items-center text-primary-600 hover:text-primary-700 transition-colors mb-8">
            <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
            </svg>
            Back to Home
          </Link>
          
          <h1 className="text-4xl sm:text-5xl font-bold text-neutral-900 mb-8">Documentation</h1>
          
          <div className="prose prose-lg max-w-none space-y-8">
            {/* Getting Started */}
            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Getting Started</h2>
              <div className="space-y-4">
                <div className="p-4 bg-pastel-sky/10 rounded-lg">
                  <h3 className="text-lg font-semibold text-neutral-800 mb-2">1. Install CellPilot</h3>
                  <p className="text-neutral-600">
                    Visit the Google Workspace Marketplace and search for "CellPilot". Click "Install" and grant the necessary permissions.
                  </p>
                </div>
                <div className="p-4 bg-pastel-mint/10 rounded-lg">
                  <h3 className="text-lg font-semibold text-neutral-800 mb-2">2. Open in Google Sheets</h3>
                  <p className="text-neutral-600">
                    Open any Google Sheets document. You'll find CellPilot in the Extensions menu.
                  </p>
                </div>
                <div className="p-4 bg-pastel-lavender/10 rounded-lg">
                  <h3 className="text-lg font-semibold text-neutral-800 mb-2">3. Start Using Features</h3>
                  <p className="text-neutral-600">
                    Select your data range and choose from our powerful features in the sidebar.
                  </p>
                </div>
              </div>
            </div>

            {/* Core Features */}
            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Core Features</h2>
              
              <div className="space-y-6">
                <div>
                  <h3 className="text-xl font-semibold text-neutral-800 mb-3">Remove Duplicates</h3>
                  <p className="text-neutral-600 mb-3">
                    Quickly identify and remove duplicate entries from your data.
                  </p>
                  <ul className="list-disc list-inside space-y-2 text-neutral-600">
                    <li>Select the range containing potential duplicates</li>
                    <li>Click "Remove Duplicates" in the CellPilot sidebar</li>
                    <li>Choose which columns to check for duplicates</li>
                    <li>Review and confirm the removal</li>
                  </ul>
                </div>

                <div>
                  <h3 className="text-xl font-semibold text-neutral-800 mb-3">Clean Text Data</h3>
                  <p className="text-neutral-600 mb-3">
                    Standardize and clean text data with multiple options.
                  </p>
                  <ul className="list-disc list-inside space-y-2 text-neutral-600">
                    <li>Trim extra spaces</li>
                    <li>Convert case (upper, lower, title)</li>
                    <li>Remove special characters</li>
                    <li>Fix encoding issues</li>
                  </ul>
                </div>

                <div>
                  <h3 className="text-xl font-semibold text-neutral-800 mb-3">Formula Builder</h3>
                  <p className="text-neutral-600 mb-3">
                    Generate complex formulas by describing what you need in plain language.
                  </p>
                  <ul className="list-disc list-inside space-y-2 text-neutral-600">
                    <li>Describe your calculation needs</li>
                    <li>CellPilot generates the appropriate formula</li>
                    <li>Preview the formula before applying</li>
                    <li>Learn from generated examples</li>
                  </ul>
                </div>

                <div>
                  <h3 className="text-xl font-semibold text-neutral-800 mb-3">Email Automation</h3>
                  <p className="text-neutral-600 mb-3">
                    Set up automated email alerts based on your spreadsheet data.
                  </p>
                  <ul className="list-disc list-inside space-y-2 text-neutral-600">
                    <li>Configure trigger conditions</li>
                    <li>Customize email templates</li>
                    <li>Schedule regular reports</li>
                    <li>Track email history</li>
                  </ul>
                </div>
              </div>
            </div>

            {/* Advanced Features */}
            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Advanced Features</h2>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="p-4 bg-neutral-50 rounded-lg">
                  <h3 className="font-semibold text-neutral-800 mb-2">Data Validation</h3>
                  <p className="text-sm text-neutral-600">
                    Ensure data integrity with custom validation rules and automatic error detection.
                  </p>
                </div>
                
                <div className="p-4 bg-neutral-50 rounded-lg">
                  <h3 className="font-semibold text-neutral-800 mb-2">Automatic Backups</h3>
                  <p className="text-sm text-neutral-600">
                    Every operation is preceded by an automatic backup for easy undo.
                  </p>
                </div>
                
                <div className="p-4 bg-neutral-50 rounded-lg">
                  <h3 className="font-semibold text-neutral-800 mb-2">Batch Operations</h3>
                  <p className="text-sm text-neutral-600">
                    Apply operations to multiple sheets or ranges simultaneously.
                  </p>
                </div>
                
                <div className="p-4 bg-neutral-50 rounded-lg">
                  <h3 className="font-semibold text-neutral-800 mb-2">Usage Analytics</h3>
                  <p className="text-sm text-neutral-600">
                    Track your usage patterns and optimize your workflow.
                  </p>
                </div>
              </div>
            </div>

            {/* Keyboard Shortcuts */}
            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Keyboard Shortcuts</h2>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="flex justify-between p-3 bg-neutral-50 rounded">
                  <span className="text-neutral-700">Open CellPilot</span>
                  <kbd className="px-2 py-1 bg-white rounded border border-neutral-300 text-sm font-mono">Alt+Shift+C</kbd>
                </div>
                <div className="flex justify-between p-3 bg-neutral-50 rounded">
                  <span className="text-neutral-700">Remove Duplicates</span>
                  <kbd className="px-2 py-1 bg-white rounded border border-neutral-300 text-sm font-mono">Alt+Shift+D</kbd>
                </div>
                <div className="flex justify-between p-3 bg-neutral-50 rounded">
                  <span className="text-neutral-700">Clean Text</span>
                  <kbd className="px-2 py-1 bg-white rounded border border-neutral-300 text-sm font-mono">Alt+Shift+T</kbd>
                </div>
                <div className="flex justify-between p-3 bg-neutral-50 rounded">
                  <span className="text-neutral-700">Formula Builder</span>
                  <kbd className="px-2 py-1 bg-white rounded border border-neutral-300 text-sm font-mono">Alt+Shift+F</kbd>
                </div>
              </div>
            </div>

            {/* Troubleshooting */}
            <div className="glass-card rounded-2xl p-8 space-y-6">
              <h2 className="text-2xl font-semibold text-neutral-900">Troubleshooting</h2>
              
              <div className="space-y-4">
                <details className="group">
                  <summary className="cursor-pointer font-semibold text-neutral-800 hover:text-primary-600">
                    CellPilot menu doesn't appear
                  </summary>
                  <p className="mt-2 text-neutral-600 pl-4">
                    Refresh your browser and ensure you've granted all necessary permissions. Try reinstalling the add-on if the issue persists.
                  </p>
                </details>
                
                <details className="group">
                  <summary className="cursor-pointer font-semibold text-neutral-800 hover:text-primary-600">
                    Operations are running slowly
                  </summary>
                  <p className="mt-2 text-neutral-600 pl-4">
                    Large datasets may take time to process. Consider breaking your data into smaller chunks or upgrading to a paid plan for faster processing.
                  </p>
                </details>
                
                <details className="group">
                  <summary className="cursor-pointer font-semibold text-neutral-800 hover:text-primary-600">
                    Formula builder returns errors
                  </summary>
                  <p className="mt-2 text-neutral-600 pl-4">
                    Ensure your description is clear and specific. Check that referenced cells and ranges exist in your sheet.
                  </p>
                </details>
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>
  )
}