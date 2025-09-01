#!/bin/bash

# Deploy CellPilot as a test add-on for beta users

echo "Deploying CellPilot as test add-on..."

# Create a new version
clasp version "Beta Test Add-on v1"

# Deploy as add-on (using the HEAD deployment which is configured for add-on)
DEPLOYMENT_ID=$(clasp deployments | grep "@HEAD" | awk '{print $2}')

echo ""
echo "========================================="
echo "CellPilot Beta Test Add-on Setup"
echo "========================================="
echo ""
echo "To install the beta version:"
echo ""
echo "1. Open this link in your browser:"
echo "   https://script.google.com/d/1EZDAGoLY8UEMdbfKTZO-AQ7pkiPe-n-zrz3Rw0ec6VBBH5MdC43Avx0O/edit"
echo ""
echo "2. Click 'Deploy' > 'Test deployments'"
echo ""
echo "3. Under 'Application(s): Sheets', click 'Install'"
echo ""
echo "4. Select the Google account to test with"
echo ""
echo "5. Open any Google Sheet with that account"
echo ""
echo "6. CellPilot will appear in the Extensions menu"
echo ""
echo "========================================="
echo ""
echo "For multiple beta testers:"
echo "1. In the Apps Script editor, click 'Deploy' > 'Test deployments'"
echo "2. Click 'Select testers' and add their email addresses"
echo "3. Share the installation link with them"
echo ""