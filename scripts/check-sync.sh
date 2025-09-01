#!/bin/bash

# CellPilot Synchronization Checker
# Ensures all proxy functions are synchronized across files

echo "🔍 CellPilot Sync Checker"
echo "========================="
echo ""

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Files to check
LIBRARY_FILE="apps-script/main-library/src/Library.js"
INSTALLER_FILE="apps-script/beta-installer/Code.gs"
TEST_PROXY_FILE="test-sheet-proxy.js"

# Extract function names from each file
echo "📋 Extracting functions from files..."

# Get exported functions from Library.js (var functionName = functionName pattern)
LIBRARY_EXPORTS=$(grep -E "^var \w+ = \w+ \|\|" "$LIBRARY_FILE" | sed -E 's/var ([a-zA-Z0-9_]+) =.*/\1/' | sort)

# Get proxy functions from beta installer (function name() { return CellPilot pattern)
INSTALLER_FUNCTIONS=$(grep -E "^function \w+\(\)" "$INSTALLER_FILE" | grep "return CellPilot\." | sed -E 's/function ([a-zA-Z0-9_]+).*/\1/' | sort)

# Get proxy functions from test-sheet-proxy.js
PROXY_FUNCTIONS=$(grep -E "^function \w+\(" "$TEST_PROXY_FILE" | sed -E 's/function ([a-zA-Z0-9_]+).*/\1/' | sort)

# Count functions
LIBRARY_COUNT=$(echo "$LIBRARY_EXPORTS" | wc -l)
INSTALLER_COUNT=$(echo "$INSTALLER_FUNCTIONS" | wc -l)
PROXY_COUNT=$(echo "$PROXY_FUNCTIONS" | wc -l)

echo ""
echo "📊 Function Counts:"
echo "  Library Exports:  $LIBRARY_COUNT"
echo "  Beta Installer:   $INSTALLER_COUNT"
echo "  Test Proxy:       $PROXY_COUNT"
echo ""

# Check for missing functions
echo "🔍 Checking for discrepancies..."
echo ""

# Functions in test proxy but not in installer
MISSING_IN_INSTALLER=$(comm -23 <(echo "$PROXY_FUNCTIONS") <(echo "$INSTALLER_FUNCTIONS"))
if [ ! -z "$MISSING_IN_INSTALLER" ]; then
    echo -e "${RED}❌ Functions in test-proxy.js but NOT in beta installer:${NC}"
    echo "$MISSING_IN_INSTALLER" | while read func; do
        echo "   - $func"
    done
    echo ""
fi

# Functions in installer but not in test proxy
MISSING_IN_PROXY=$(comm -23 <(echo "$INSTALLER_FUNCTIONS") <(echo "$PROXY_FUNCTIONS"))
if [ ! -z "$MISSING_IN_PROXY" ]; then
    echo -e "${YELLOW}⚠️  Functions in beta installer but NOT in test-proxy.js:${NC}"
    echo "$MISSING_IN_PROXY" | while read func; do
        echo "   - $func"
    done
    echo ""
fi

# Check if all three are in sync
if [ "$INSTALLER_FUNCTIONS" = "$PROXY_FUNCTIONS" ]; then
    echo -e "${GREEN}✅ Beta installer and test proxy are in sync!${NC}"
else
    echo -e "${RED}❌ Beta installer and test proxy are NOT in sync${NC}"
fi

echo ""
echo "📝 Next Steps:"
if [ ! -z "$MISSING_IN_INSTALLER" ] || [ ! -z "$MISSING_IN_PROXY" ]; then
    echo "1. Add missing functions to ensure consistency"
    echo "2. Run: cd apps-script && npm run push:installer"
    echo "3. Commit changes: git add -A && git commit -m 'Sync proxy functions'"
else
    echo "Everything looks good! 🎉"
fi

echo ""
echo "💡 Tip: Run this script before deploying to ensure everything is synchronized"