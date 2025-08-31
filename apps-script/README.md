# CellPilot Apps Script Projects

This directory contains all Google Apps Script projects for CellPilot.

## Structure

```
apps-script/
├── main-library/        # Main CellPilot library (users reference this)
│   ├── src/            # All source files
│   ├── *.html          # HTML templates
│   ├── appsscript.json # Manifest
│   └── .clasp.json     # Script ID: 1EZDAGoLY8UEMdbfKTZO...
│
├── beta-installer/      # Web app for beta installation
│   ├── Code.gs         # Installer logic
│   ├── installer.html  # UI
│   ├── appsscript.json # Manifest
│   └── .clasp.json     # Script ID: 1PAypZbjwbnMf22...
│
└── package.json        # NPM scripts for easy management
```

## Quick Commands

From the `apps-script` directory:

### Push Changes
```bash
npm run push:library      # Push main library changes
npm run push:installer    # Push installer changes
npm run push:all         # Push both
```

### Deploy New Versions
```bash
npm run deploy:library "Description"    # Deploy library
npm run deploy:installer "Description"  # Deploy installer
```

### Open in Browser
```bash
npm run open:library     # Open main library in Apps Script editor
npm run open:installer   # Open installer in Apps Script editor
```

### Check Status
```bash
npm run status:library   # Check library status
npm run status:installer # Check installer status
```

### View Logs
```bash
npm run logs:library     # View library execution logs
npm run logs:installer   # View installer logs
```

## Deployment Info

### Main Library
- **Script ID**: `1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O`
- **Current Version**: 10
- **Purpose**: Core CellPilot functionality that users reference

### Beta Installer
- **Script ID**: `1PAypZbjwbnMf22dyHXItxbyRIhCZgIME4JtR8REpzk_PfvH17sg-h33f`
- **Current Deployment**: v1.5
- **Web App URL**: See BETA_INSTALLER_WEB_APP.txt
- **Purpose**: Web interface for beta users to get installation code

## Workflow

1. Make changes to files in respective directories
2. Use npm scripts to push changes
3. Deploy when ready for new version
4. Update site/components/BetaAccessCard.tsx with new deployment URLs if needed

## Important Notes

- Each project has its own `.clasp.json` with correct script ID
- Don't mix files between projects
- Always push before deploying
- Deployments create new URLs for web apps