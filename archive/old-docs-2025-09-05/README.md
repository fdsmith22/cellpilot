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
