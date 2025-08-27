# CellPilot - Google Sheets Add-on

A powerful Google Sheets add-on for spreadsheet automation, data cleaning, and formula generation.

## Project Structure

```
cellpilot/
├── src/                    # JavaScript/Google Apps Script files
│   ├── Code.js            # Main entry point & hybrid UI controller
│   ├── Config.js          # Configuration, settings & pricing tiers
│   ├── Utils.js           # Utility functions & error handling
│   ├── DataCleaner.js     # Data cleaning operations
│   ├── FormulaBuilder.js  # Natural language formula generation
│   └── UIComponents.js    # CardService UI components (legacy)
│
├── html/                   # HTML templates for sidebars
│   ├── DuplicateRemovalTemplate.html
│   └── FormulaBuilderTemore toolsmplate.html
│
├── site/                   # Marketing website (Next.js)
│   ├── app/               # App router pages
│   ├── components/        # React components
│   ├── lib/               # Utility libraries
│   ├── public/            # Static assets
│   └── package.json       # Website dependencies
│
├── scripts/               # Development scripts
│   └── dev.sh            # Unified development helper
│
├── appsscript.json        # Google Apps Script manifest
├── .clasp.json            # Clasp configuration
├── .gitignore             # Version control ignore file
└── README.md              # This file
```

## Architecture

CellPilot uses a **hybrid architecture** supporting both:

1. **CardService UI** - For Google Workspace Marketplace distribution
2. **HTML Sidebars** - For development and independent installation

## Development Workflow

```bash
# Install clasp globally (one-time)
npm install -g @google/clasp

# Login to Google (one-time)
clasp login

# Make changes to files
# ...

# Push changes to Google Apps Script
clasp push

# Open in browser for testing
clasp open
```

## Features

### Data Cleaning
- **Duplicate Removal** - Smart duplicate detection with fuzzy matching
- **Text Standardization** - Case formatting, whitespace cleanup
- **Date Formatting** - (Coming soon)

### Formula Builder
- Natural language to Google Sheets formulas
- Supports SUM, AVERAGE, COUNT, VLOOKUP, IF, and more
- Context-aware formula suggestions

### Automation
- Email alerts (Premium)
- Calendar integration (Premium)
- Batch processing

## Pricing Tiers

- **Free**: 25 operations/month, basic features
- **Starter** (£5.99/month): 500 operations, formula builder
- **Professional** (£11.99/month): Unlimited operations, all features
- **Business** (£19.99/month): Team features, priority support

## Installation Methods

### For Development/Testing
1. Open any Google Sheet
2. Extensions > Apps Script
3. Copy script ID: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
4. Add as library or copy files
5. Run `onOpen()` function
6. Refresh sheet to see CellPilot menu

### For End Users (Coming Soon)
- Google Workspace Marketplace
- Direct installation from cellpilot.io

## Testing

After pushing changes:
1. Open a Google Sheet
2. Look for "CellPilot" menu
3. Test features:
   - CellPilot > Open CellPilot (main sidebar)
   - Data Cleaning > Remove Duplicates
   - Formula Builder > Natural Language Builder

## Support

- Website: https://www.cellpilot.io
- Documentation: Coming soon
- Support: support@cellpilot.io

## License

Proprietary - All rights reserved