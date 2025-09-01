# CellPilot Project Structure

## Overview
CellPilot is a Google Sheets add-on with a Next.js web application for user management.

## Directory Structure

```
cellpilot/
├── apps-script/           # Google Apps Script code
│   ├── main-library/     # Core CellPilot library
│   └── beta-installer/   # Web app for installation
├── site/                 # Next.js web application
│   ├── app/             # App router pages
│   ├── components/      # React components
│   ├── lib/            # Utilities and helpers
│   └── supabase/       # Database migrations
├── docs/                # Documentation
├── archive/            # Archived/outdated files
└── test-sheet-proxy.js # Proxy functions for testing
```

## Key Files

- `apps-script/main-library/src/Code.js` - Main Apps Script entry point
- `apps-script/beta-installer/Code.gs` - Beta installer web app
- `site/app/dashboard/page.tsx` - User dashboard
- `test-sheet-proxy.js` - Complete proxy functions for user installation

## Development Workflow

1. Make changes to Apps Script code in `apps-script/` folder
2. Push to Google Apps Script using npm scripts
3. Test in Google Sheets
4. Update web app as needed in `site/` folder
5. Deploy web changes to Vercel
