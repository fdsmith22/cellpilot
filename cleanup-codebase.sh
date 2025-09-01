#!/bin/bash

echo "🧹 CellPilot Codebase Cleanup Script"
echo "===================================="

# Create archive folder for outdated docs
echo "📁 Creating archive folder for outdated documentation..."
mkdir -p archive/beta-docs
mkdir -p archive/old-scripts

# Move outdated Beta documentation to archive
echo "📦 Archiving outdated Beta documentation..."
mv BETA_DEPLOYMENT_GUIDE.md archive/beta-docs/ 2>/dev/null
mv BETA_INSTALLATION_SUMMARY.md archive/beta-docs/ 2>/dev/null
mv BETA_INSTALLER.gs archive/beta-docs/ 2>/dev/null
mv BETA_INSTALLER_WEB_APP.txt archive/beta-docs/ 2>/dev/null
mv BETA_LAUNCH_CHECKLIST.md archive/beta-docs/ 2>/dev/null
mv BETA_READINESS_CHECKLIST.md archive/beta-docs/ 2>/dev/null
mv BETA_USER_CODE.gs archive/beta-docs/ 2>/dev/null
mv DEPLOYMENT_GUIDE.md archive/beta-docs/ 2>/dev/null
mv INTEGRATION_SUMMARY.md archive/beta-docs/ 2>/dev/null
mv LIBRARY_DEPLOYMENT_INFO.txt archive/beta-docs/ 2>/dev/null

# Move outdated testing docs to archive
echo "📦 Archiving outdated testing documentation..."
mv TESTING.md archive/beta-docs/ 2>/dev/null
mv TESTING_CHECKLIST.md archive/beta-docs/ 2>/dev/null
mv TESTING_GUIDE.md archive/beta-docs/ 2>/dev/null

# Remove duplicate HTML templates from root (keep only in apps-script folder)
echo "🗑️  Removing duplicate HTML templates from root..."
rm -f AdvancedRestructuringTemplate.html
rm -f AutomationTemplate.html
rm -f ConditionalFormattingWizardTemplate.html
rm -f CrossSheetFormulaBuilderTemplate.html
rm -f DataPipelineTemplate.html
rm -f DataValidationGeneratorTemplate.html
rm -f DuplicateRemovalTemplate.html
rm -f FormulaBuilderTemplate.html
rm -f HelpFeedbackTemplate.html
rm -f IndustryTemplatesTemplate.html
rm -f MLEngine.html
rm -f MLEngineLoader.html
rm -f MultiTabRelationshipMapperTemplate.html
rm -f OnboardingTemplate.html
rm -f PivotTableAssistantTemplate.html
rm -f SettingsTemplate.html
rm -f SharedStyles.html
rm -f SmartFormulaAssistantTemplate.html
rm -f SmartFormulaDebuggerTemplate.html
rm -f TableizeTemplate.html
rm -f UpgradeTemplate.html

# Remove duplicate src folder from root (keep only in apps-script/main-library)
echo "🗑️  Removing duplicate src folder from root..."
rm -rf src/

# Remove duplicate html folder (already in apps-script)
echo "🗑️  Removing duplicate html folder..."
rm -rf html/

# Move old scripts to archive
echo "📦 Archiving old scripts..."
mv cleanup-test-user.sh archive/old-scripts/ 2>/dev/null
mv deploy-test-addon.sh archive/old-scripts/ 2>/dev/null
mv replace_console_logs.sh archive/old-scripts/ 2>/dev/null
mv update_styles.sh archive/old-scripts/ 2>/dev/null

# Remove duplicate appsscript.json from root
echo "🗑️  Removing duplicate appsscript.json from root..."
rm -f appsscript.json

# Move test-sheet-proxy backup
echo "📦 Archiving test-sheet-proxy backup..."
mv test-sheet-proxy.js.backup archive/ 2>/dev/null

# Create updated documentation structure
echo "📝 Creating updated documentation structure..."
mkdir -p docs

# Create new README for apps-script folder if needed
cat > apps-script/README.md << 'EOF'
# CellPilot Apps Script Projects

This folder contains all Google Apps Script related code for CellPilot.

## Structure

### `/main-library`
The main CellPilot library code that gets deployed to Google Apps Script.
- Contains all core functionality
- HTML templates for UI components
- Deployed as a library for users to reference

### `/beta-installer`
Web app that helps users install CellPilot in their Google Sheets.
- Provides installation instructions
- Generates proxy code for users
- Checks beta access

## Development

Use npm scripts from the root `apps-script` folder:

```bash
# Push main library
npm run push:library

# Push beta installer
npm run push:installer

# Pull changes
npm run pull:library
npm run pull:installer
```

## Deployment

The main library is deployed as version-controlled releases.
The beta installer is deployed as a web app with public access.
EOF

# Create main project documentation
cat > docs/PROJECT_STRUCTURE.md << 'EOF'
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
EOF

# Update main README
cat > README.md << 'EOF'
# CellPilot

Advanced data manipulation and automation tool for Google Sheets.

## 🚀 Features

- **Smart Table Parser**: Convert unstructured data into organized columns
- **Advanced Data Cleaning**: Remove duplicates, standardize text, format dates
- **Formula Builder**: Create complex formulas using natural language
- **Industry Templates**: Pre-built templates for various industries
- **ML-Powered Features**: Smart suggestions and pattern recognition
- **Cross-Sheet Operations**: Work with data across multiple sheets
- **Data Pipeline Manager**: Import/export data with transformations

## 📦 Installation

### For Beta Users

1. Sign up at [cellpilot.io](https://www.cellpilot.io)
2. Activate beta access in your dashboard
3. Click "Install CellPilot" 
4. Follow the installation guide

### For Developers

See [docs/PROJECT_STRUCTURE.md](docs/PROJECT_STRUCTURE.md) for development setup.

## 🛠️ Technology Stack

- **Google Apps Script**: Core add-on functionality
- **Next.js 14**: Web application and user dashboard
- **Supabase**: Authentication and database
- **Tailwind CSS**: Styling
- **TypeScript**: Type safety

## 📁 Project Structure

```
cellpilot/
├── apps-script/        # Google Apps Script code
├── site/              # Next.js web application  
├── docs/              # Documentation
└── test-sheet-proxy.js # Testing proxy functions
```

## 🔧 Development

### Apps Script Development

```bash
cd apps-script
npm install

# Push to Google Apps Script
npm run push:library
npm run push:installer
```

### Web Development

```bash
cd site
npm install
npm run dev
```

## 📄 License

Proprietary - All rights reserved

## 🤝 Support

For support, email support@cellpilot.io
EOF

echo ""
echo "✅ Cleanup complete!"
echo ""
echo "Summary of changes:"
echo "- Archived outdated Beta documentation to archive/beta-docs/"
echo "- Removed duplicate HTML templates from root (kept in apps-script/main-library/)"
echo "- Removed duplicate src/ folder from root"
echo "- Archived old scripts to archive/old-scripts/"
echo "- Created updated documentation in docs/"
echo "- Updated README files"
echo ""
echo "The codebase is now clean and organized!"