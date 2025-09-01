# CellPilot Development Workflow Guide

## 🎯 Overview
This guide ensures all components stay synchronized when making changes to CellPilot.

## 📦 Project Components

1. **Main Library** (`apps-script/main-library/`) - Core functionality
2. **Beta Installer** (`apps-script/beta-installer/`) - Installation web app
3. **Test Proxy** (`test-sheet-proxy.js`) - User-facing proxy functions
4. **Website** (`site/`) - Next.js dashboard
5. **Database** (Supabase) - User data and authentication

## 🔄 Development Workflows

### 1️⃣ Adding a New Feature to CellPilot


#### Step 1: Implement in Main Library
```bash
cd /home/freddy/cellpilot/apps-script/main-library

# Edit the appropriate file
nano src/[FeatureName].js  # or create new file

# If adding new UI, create HTML template
nano [FeatureName]Template.html
```

#### Step 2: Export the Function
```javascript
// In src/Library.js, add export at the bottom:
var showNewFeature = showNewFeature || function() { return showNewFeature(); };
```

#### Step 3: Add Proxy Function to Beta Installer
```bash
cd /home/freddy/cellpilot/apps-script/beta-installer

# Edit Code.gs - add in the appropriate section:
nano Code.gs
```

Add the proxy function:
```javascript
function showNewFeature() { return CellPilot.showNewFeature(); }
```

#### Step 4: Add to Test Proxy
```bash
cd /home/freddy/cellpilot

# Edit test-sheet-proxy.js - add the same proxy function
nano test-sheet-proxy.js
```

#### Step 5: Push Changes
```bash
# Push main library
cd apps-script
npm run push:library

# Push beta installer  
npm run push:installer

# Commit to Git
cd /home/freddy/cellpilot
git add -A
git commit -m "Add [feature name] functionality"
git push
```

#### Step 6: Deploy New Versions
```bash
# Deploy new library version (if needed)
cd apps-script/main-library
npx clasp deploy -d "Version X - Added [feature]"

# Deploy new beta installer (if proxy changed)
cd ../beta-installer
npx clasp deploy -d "v1.X - Added [feature] proxy"
```

#### Step 7: Update Website
```bash
cd /home/freddy/cellpilot/site

# Update BetaAccessCard.tsx with new deployment URL
nano components/BetaAccessCard.tsx

# Update the installer URL to new deployment ID
```

### 2️⃣ Modifying Existing Feature

#### Quick Process:
```bash
# 1. Make changes in main library
cd apps-script/main-library
# Edit files...

# 2. Push to Apps Script
cd ..
npm run push:library

# 3. Test in Google Sheets
# Open test sheet with library reference

# 4. If working, commit
cd /home/freddy/cellpilot
git add -A
git commit -m "Update [feature]: [change description]"
git push
```

### 3️⃣ Database Changes (Supabase)

#### Adding Migration:
```bash
cd /home/freddy/cellpilot/site

# Create migration
npx supabase migration new [migration_name]

# Edit the migration file
nano supabase/migrations/[timestamp]_[migration_name].sql

# Apply locally
npx supabase db reset

# Push to production
npx supabase db push
```

#### Updating RLS Policies:
```bash
# Always test locally first
npx supabase db reset --local

# Then push
npx supabase db push
```

### 4️⃣ Website Updates

#### Development:
```bash
cd site
npm run dev
# Make changes, test locally

# Deploy to production
git add -A
git commit -m "Update website: [changes]"
git push
# Vercel auto-deploys from GitHub
```

## 📋 Synchronization Checklist

When adding a new feature, ensure ALL of these are updated:

### Apps Script Components:
- [ ] Main library implementation (`apps-script/main-library/src/`)
- [ ] Library exports (`src/Library.js`)
- [ ] Beta installer proxy (`apps-script/beta-installer/Code.gs`)
- [ ] Test proxy file (`test-sheet-proxy.js`)
- [ ] HTML templates if needed

### Deployments:
- [ ] Push main library (`npm run push:library`)
- [ ] Push beta installer (`npm run push:installer`)
- [ ] Create new deployments if needed
- [ ] Update installer URL in `site/components/BetaAccessCard.tsx`

### Documentation:
- [ ] Update function list in docs if needed
- [ ] Add JSDoc comments to new functions
- [ ] Update README if major feature

### Git:
- [ ] Commit all changes together
- [ ] Use descriptive commit message
- [ ] Push to GitHub

## 🛠️ Common Commands Reference

### Apps Script (from `/apps-script` folder):
```bash
# Push changes
npm run push:library    # Push main library
npm run push:installer  # Push beta installer

# Pull changes
npm run pull:library    # Pull main library
npm run pull:installer  # Pull beta installer

# Deploy new versions
cd main-library && npx clasp deploy -d "Description"
cd beta-installer && npx clasp deploy -d "Description"

# View deployments
npx clasp deployments
```

### Supabase (from `/site` folder):
```bash
# Local development
npx supabase start      # Start local Supabase
npx supabase stop       # Stop local Supabase

# Database
npx supabase db reset   # Reset local database
npx supabase db push    # Push migrations to production

# Check status
npx supabase status
```

### Git:
```bash
# Standard flow
git add -A
git commit -m "Type: Description"
git push

# Check what changed
git status
git diff
```

## 🔍 Testing Workflow

### 1. Test Library Changes:
```bash
# After pushing library changes
# Open: https://sheets.new
# Extensions > Apps Script
# Add library with HEAD version
# Paste test-sheet-proxy.js
# Test the feature
```

### 2. Test Beta Installer:
```bash
# After deploying new version
# Open the deployment URL
# Go through installation process
# Verify proxy functions work
```

### 3. Test Website:
```bash
cd site
npm run dev
# Test at http://localhost:3000
```

## ⚠️ Important Notes

### Version Control:
- Main library version only matters for releases
- Beta installer should always use library version 11 or HEAD
- Website deployment URL must match latest beta installer

### Clasp Configurations:
- Each Apps Script project has its own `.clasp.json`
- Don't mix them up - always be in the right directory
- Use npm scripts from `/apps-script` folder

### Order of Operations:
1. Always implement in main library first
2. Test locally before pushing
3. Update all proxy locations
4. Deploy new versions
5. Update website references
6. Commit everything together

## 🚨 Troubleshooting

### "Function not found" in sidebar:
1. Check function exists in main library
2. Verify it's exported in Library.js
3. Confirm proxy exists in beta installer Code.gs
4. Ensure names match exactly (case-sensitive)

### Changes not appearing:
1. Hard refresh Google Sheets (Ctrl+Shift+R)
2. Check you pushed to Apps Script
3. Verify correct deployment version
4. Clear browser cache if needed

### Database errors:
1. Check Supabase status: `npx supabase status`
2. Verify RLS policies
3. Check migration order
4. Review Supabase logs

## 📝 Example: Adding "Data Export" Feature

```bash
# 1. Implement in main library
cd apps-script/main-library
# Create src/DataExporter.js
# Add showDataExporter function

# 2. Export it
# Edit src/Library.js
# Add: var showDataExporter = showDataExporter || function() { return showDataExporter(); };

# 3. Add proxy to beta installer
cd ../beta-installer
# Edit Code.gs
# Add: function showDataExporter() { return CellPilot.showDataExporter(); }

# 4. Add to test proxy
cd ../..
# Edit test-sheet-proxy.js
# Add the same proxy function

# 5. Push everything
cd apps-script
npm run push:library
npm run push:installer

# 6. Test in Google Sheets
# Open test sheet, verify it works

# 7. Deploy new versions
cd main-library
npx clasp deploy -d "v12 - Added Data Export"

cd ../beta-installer  
npx clasp deploy -d "v1.9 - Added Data Export proxy"
# Copy new deployment ID

# 8. Update website
cd ../../site
# Edit components/BetaAccessCard.tsx
# Update URL to new deployment

# 9. Commit everything
cd ..
git add -A
git commit -m "Add Data Export feature

- Implemented DataExporter in main library
- Added proxy functions to beta installer and test proxy
- Deployed as library v12 and installer v1.9
- Updated website to use new installer"
git push
```

## 🔐 Security Reminders

1. Never commit `.clasp.json` files (they contain project IDs)
2. Keep Supabase keys in `.env.local` only
3. Don't expose library source code in beta installer
4. Test RLS policies thoroughly
5. Review code before pushing to production

---

*Last updated: September 1, 2025*