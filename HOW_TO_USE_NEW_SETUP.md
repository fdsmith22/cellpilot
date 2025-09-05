# 🚀 How to Use the New CellPilot Development Setup

## Quick Reference - Essential Commands

### 📊 Check Project Status
```bash
./scripts/status.sh
```
Shows real-time metrics: file counts, git status, dependencies, CellM8 stats

### 📖 Access Documentation
```bash
cat MASTER_DOCUMENTATION.md | less
# Or search specific sections:
grep -A20 "CellM8" MASTER_DOCUMENTATION.md
```

### 🤖 Generate AI Context
```bash
codebase-digest --path apps-script/main-library/src --output ai-context.txt
```
Creates AI-friendly summary of your codebase

### 🚀 Deploy Code
```bash
cd apps-script
npm run push:library     # Push changes (no version)
npm run deploy:library   # Deploy new version
npm run logs:library     # View execution logs
```

---

## Tool-Specific Usage

### 1. **JSDoc - Generate Documentation**
```bash
cd apps-script
# Generate HTML documentation
npx jsdoc main-library/src -r -d ../docs/api

# Generate markdown documentation
npx jsdoc2md main-library/src/*.js > ../docs/API.md
```

### 2. **ESLint - Code Quality**
```bash
cd apps-script
# Check all files
npx eslint main-library/src

# Auto-fix issues
npx eslint main-library/src --fix

# Check specific file
npx eslint main-library/src/CellM8.js
```

### 3. **Prettier - Code Formatting**
```bash
cd apps-script
# Format all files
npx prettier --write main-library/src/*.js

# Check formatting without changing
npx prettier --check main-library/src/*.js
```

### 4. **Dependency Cruiser - Visualize Dependencies**
```bash
cd apps-script
# Generate dependency graph
npx dependency-cruiser main-library/src --output-type dot | dot -T svg > dependency-graph.svg

# Check for circular dependencies
npx dependency-cruiser main-library/src --config .dependency-cruiser.js
```

### 5. **Husky - Git Hooks (Auto-runs)**
```bash
# Already configured to run on commit
# Manually run pre-commit checks:
cd apps-script
npm run lint
npm run format:check
```

---

## PyCharm Integration

### Run Configurations
1. **Deploy CellPilot**: Click Run → Deploy CellPilot
2. **Project Status**: Click Run → Project Status

### File Watchers
- **Prettier**: Auto-formats on save (already configured)
- **ESLint**: Shows errors in editor (optional, disabled by default)

### Shortcuts
- `Ctrl+Shift+F10`: Run current configuration
- `Alt+Shift+F`: Format current file
- `Ctrl+Alt+L`: Reformat code

---

## Development Workflow

### 1. **Starting Work**
```bash
# Check project status
./scripts/status.sh

# Pull latest changes
git pull

# Check documentation
cat MASTER_DOCUMENTATION.md | grep -A10 "section_you_need"
```

### 2. **Making Changes**
```bash
# 1. Edit files in PyCharm
# 2. Test locally
cd apps-script && npm run push:library

# 3. Open Apps Script editor
npm run open:library

# 4. Test in spreadsheet
# 5. Check logs
npm run logs:library
```

### 3. **Before Committing**
```bash
# Format code
cd apps-script
npx prettier --write main-library/src/*.js

# Check for issues
npx eslint main-library/src

# Run status check
./scripts/status.sh

# Commit (Husky will auto-check)
git add .
git commit -m "feat: your change description"
```

### 4. **Deploying**
```bash
cd apps-script
# Deploy new version with description
npm run deploy:library "Version X.X - Description"

# Verify deployment
npm run status:library
```

---

## Important Files & Locations

### 📁 Core Documentation
- `MASTER_DOCUMENTATION.md` - Single source of truth (600+ lines)
- `.claude/claude.md` - AI assistant context
- `archive/old-docs-2025-09-05/` - Old documentation (archived)

### 📁 Scripts
- `scripts/status.sh` - Project health dashboard
- `scripts/dev.sh` - Development helper
- `scripts/check-sync.sh` - Sync verification

### 📁 Configuration
- `apps-script/package.json` - NPM scripts and dependencies
- `apps-script/main-library/.clasp.json` - CLASP config
- `.idea/` - PyCharm settings
- `pyproject.toml` - Python/PyCharm project config

---

## Common Tasks

### Find a Function
```bash
# Search in all JS files
grep -r "functionName" apps-script/main-library/src

# Find with context
grep -B2 -A2 "analyzeDataIntelligently" apps-script/main-library/src/CellM8.js
```

### Update Documentation
```bash
# Edit the master doc
nano MASTER_DOCUMENTATION.md
# or in PyCharm: Ctrl+Shift+N → MASTER_DOCUMENTATION.md
```

### Generate Dependency Report
```bash
cd apps-script
npx dependency-cruiser main-library/src --output-type html > dependency-report.html
open dependency-report.html
```

### Check What Changed
```bash
git status
git diff apps-script/main-library/src/Code.js
```

---

## Troubleshooting

### If status.sh shows errors
```bash
chmod +x scripts/status.sh
./scripts/status.sh
```

### If ESLint complains
```bash
cd apps-script
npm install  # Reinstall dependencies
npx eslint --init  # Reconfigure if needed
```

### If deployment fails
```bash
cd apps-script
clasp login  # Re-authenticate
npm run status:library  # Check connection
```

### If PyCharm doesn't recognize files
1. File → Invalidate Caches and Restart
2. Check `.idea/` folder exists
3. Mark `apps-script/main-library/src` as Sources Root

---

## Best Practices

### ✅ DO
- Always check `MASTER_DOCUMENTATION.md` before major changes
- Run `./scripts/status.sh` regularly
- Use `npm run push:library` for testing (not deploy)
- Test in Apps Script editor before deploying
- Keep backwards compatibility

### ❌ DON'T
- Create new .md files (update MASTER_DOCUMENTATION.md)
- Deploy without testing
- Ignore ESLint warnings
- Commit without checking status
- Change core architecture without planning

---

## Quick Tips

1. **Speed up development**: Use `npm run push:library` instead of deploy
2. **Debug faster**: Keep Apps Script logs open in separate tab
3. **Find code quickly**: Use PyCharm's `Ctrl+Shift+N` for files
4. **Track changes**: Run `git status` frequently
5. **Stay organized**: Update todos in MASTER_DOCUMENTATION.md

---

## Need Help?

1. Check `MASTER_DOCUMENTATION.md` first
2. Run `./scripts/status.sh` to diagnose issues
3. Review git history: `git log --oneline -10`
4. Check Apps Script logs: `npm run logs:library`

---

*Remember: MASTER_DOCUMENTATION.md is your single source of truth!*