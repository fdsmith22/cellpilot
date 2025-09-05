#!/bin/bash
# CellPilot Project Status Dashboard
# Shows project health and metrics at a glance

echo "=================================================="
echo "📊 CellPilot Status Dashboard"
echo "=================================================="
echo ""

# Project Statistics
echo "📁 PROJECT STATISTICS:"
echo "----------------------"
echo "Total JS Files: $(find apps-script/main-library/src -name "*.js" 2>/dev/null | wc -l)"
echo "Total HTML Templates: $(find apps-script/main-library -name "*.html" 2>/dev/null | wc -l)"
echo "Total Functions: $(grep -r "^function " apps-script/main-library/src 2>/dev/null | wc -l)"
echo "Total Lines of Code: $(find apps-script/main-library/src -name "*.js" -exec wc -l {} \; 2>/dev/null | awk '{sum+=$1} END {print sum}')"
echo ""

# Git Status
echo "🔄 GIT STATUS:"
echo "--------------"
echo "Current Branch: $(git branch --show-current 2>/dev/null || echo 'N/A')"
echo "Modified Files: $(git status --porcelain 2>/dev/null | wc -l)"
echo "Uncommitted Changes: $(git diff --name-only 2>/dev/null | wc -l)"
echo "Last Commit: $(git log -1 --pretty=format:'%h - %s (%cr)' 2>/dev/null || echo 'N/A')"
echo ""

# Documentation
echo "📚 DOCUMENTATION:"
echo "-----------------"
echo "Master Doc: $(if [ -f "MASTER_DOCUMENTATION.md" ]; then echo "✅ Present"; else echo "❌ Missing"; fi)"
echo "Claude Config: $(if [ -f ".claude/claude.md" ]; then echo "✅ Present"; else echo "❌ Missing"; fi)"
echo "Total MD Files: $(find . -name "*.md" | wc -l)"
echo ""

# Dependencies
echo "📦 DEPENDENCIES:"
echo "----------------"
if [ -f "apps-script/package.json" ]; then
    echo "Dev Dependencies: $(grep -c '".*":' apps-script/package.json)"
    echo "ESLint: $(if grep -q "eslint" apps-script/package.json; then echo "✅ Installed"; else echo "❌ Not installed"; fi)"
    echo "Prettier: $(if grep -q "prettier" apps-script/package.json; then echo "✅ Installed"; else echo "❌ Not installed"; fi)"
    echo "JSDoc: $(if grep -q "jsdoc" apps-script/package.json; then echo "✅ Installed"; else echo "❌ Not installed"; fi)"
    echo "Dependency Cruiser: $(if grep -q "dependency-cruiser" apps-script/package.json; then echo "✅ Installed"; else echo "❌ Not installed"; fi)"
    echo "Husky: $(if grep -q "husky" apps-script/package.json; then echo "✅ Installed"; else echo "❌ Not installed"; fi)"
else
    echo "No package.json found"
fi
echo ""

# CellM8 Status
echo "🤖 CELLM8 STATUS:"
echo "-----------------"
if [ -f "apps-script/main-library/src/CellM8.js" ]; then
    echo "CellM8.js Size: $(wc -l apps-script/main-library/src/CellM8.js | awk '{print $1}') lines"
    echo "Functions: $(grep -c "^[[:space:]]*[a-zA-Z_][a-zA-Z0-9_]*:" apps-script/main-library/src/CellM8.js 2>/dev/null || echo '0')"
else
    echo "CellM8.js not found"
fi
echo ""

# Deployment Status
echo "🚀 DEPLOYMENT:"
echo "--------------"
echo "CLASP Config: $(if [ -f "apps-script/main-library/.clasp.json" ]; then echo "✅ Present"; else echo "❌ Missing"; fi)"
echo "Apps Script Manifest: $(if [ -f "apps-script/main-library/appsscript.json" ]; then echo "✅ Present"; else echo "❌ Missing"; fi)"
echo ""

echo "=================================================="
echo "Run 'cd apps-script && npm run push:library' to deploy"
echo "Run 'codebase-digest' to generate AI context"
echo "=================================================="