#!/bin/bash

# File: cellpilot/scripts/dev.sh
# Development helper script for the unified workspace

echo "CellPilot Development Environment"
echo "====================================="

# Function to start website development
start_site() {
    echo "Starting website development server..."
    cd site && npm run dev
}

# Function to deploy apps script
deploy_script() {
    echo "Deploying Apps Script updates..."
    clasp push
    echo "Apps Script updated"
}

# Function to check status of both projects
status() {
    echo "Project Status:"
    echo "---"
    echo "Apps Script files:"
    ls -la src/*.js html/*.html *.json 2>/dev/null | grep -v site
    echo "---"
    echo "Website status:"
    cd site && npm list --depth=0 2>/dev/null || echo "Website dependencies not installed"
    cd ..
}

# Function to setup everything
setup() {
    echo "Setting up CellPilot workspace..."
    
    # Check if clasp is configured
    if [ ! -f ".clasp.json" ]; then
        echo "ERROR: clasp not configured. Run 'clasp clone' first."
        exit 1
    fi
    
    # Setup website if not already done
    if [ ! -d "site/node_modules" ]; then
        echo "Installing website dependencies..."
        cd site && npm install && cd ..
    else
        echo "Website dependencies already installed"
    fi
    
    echo "Workspace setup complete!"
    echo "Run './scripts/dev.sh site' to start website development"
    echo "Run './scripts/dev.sh deploy' to update Apps Script"
}

# Main script logic
case "$1" in
    "site")
        start_site
        ;;
    "deploy")
        deploy_script
        ;;
    "status")
        status
        ;;
    "setup")
        setup
        ;;
    *)
        echo "Usage: $0 {site|deploy|status|setup}"
        echo ""
        echo "Commands:"
        echo "  site   - Start website development server"
        echo "  deploy - Push Apps Script updates"
        echo "  status - Check project status"
        echo "  setup  - Initialize workspace"
        echo ""
        echo "Current directory structure:"
        echo "cellpilot/"
        echo "├── src/ (Apps Script files)"
        echo "├── html/ (HTML templates)"
        echo "├── appsscript.json"
        echo "├── .clasp.json"
        echo "└── site/ (Website files)"
        exit 1
        ;;
esac