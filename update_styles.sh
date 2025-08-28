#!/bin/bash

# Update all HTML template files with consistent styling

echo "Updating template styles for consistency..."

# List of files to update
files=(
  "DataValidationGeneratorTemplate.html"
  "SmartFormulaDebuggerTemplate.html"
  "VisualFormulaBuilderTemplate.html"
  "AutomationTemplate.html"
  "DuplicateRemovalTemplate.html"
  "FormulaBuilderTemplate.html"
  "HelpFeedbackTemplate.html"
  "IndustryTemplatesTemplate.html"
  "MultiTabRelationshipMapperTemplate.html"
  "SmartFormulaAssistantTemplate.html"
  "TableizeTemplate.html"
  "UpgradeTemplate.html"
)

# Color replacements (old -> CSS variable)
declare -A colors=(
  ["#2c3e50"]="var(--gray-900)"
  ["#34495e"]="var(--gray-800)"
  ["#7f8c8d"]="var(--gray-500)"
  ["#95a5a6"]="var(--gray-400)"
  ["#bdc3c7"]="var(--gray-300)"
  ["#ecf0f1"]="var(--gray-100)"
  ["#e0e0e0"]="var(--gray-200)"
  ["#ddd"]="var(--gray-300)"
  ["#3498db"]="var(--primary-600)"
  ["#2980b9"]="var(--primary-700)"
  ["#e74c3c"]="var(--danger-600)"
  ["#c0392b"]="var(--danger-700)"
  ["#27ae60"]="var(--success-600)"
  ["#2ecc71"]="var(--success-500)"
  ["#f39c12"]="var(--warning-600)"
  ["#667eea"]="var(--primary-600)"
  ["#764ba2"]="var(--primary-700)"
)

# Emoji replacements
declare -A emojis=(
  ["📍"]="RP"
  ["🎯"]="CV"
  ["📝"]="TC"
  ["📅"]="DR"
  ["👥"]="DD"
  ["🌈"]="CS"
  ["📊"]="DB"
  ["🔢"]="CF"
  ["🔝"]="TB"
  ["✅"]="OK"
  ["❌"]="NO"
  ["🔍"]="SR"
  ["📈"]="CH"
  ["🚀"]="GO"
  ["⚡"]="ZP"
  ["💾"]="SV"
  ["📋"]="CL"
  ["🏠"]="HM"
  ["⬅️"]="BK"
  ["➡️"]="NX"
  ["⚠️"]="WN"
  ["ℹ️"]="IN"
  ["🔧"]="ST"
  ["🎨"]="TH"
  ["📁"]="FL"
  ["💡"]="ID"
  ["🔄"]="RF"
)

for file in "${files[@]}"; do
  if [ -f "/home/freddy/cellpilot/$file" ]; then
    echo "Processing $file..."
    
    # Create backup
    cp "/home/freddy/cellpilot/$file" "/home/freddy/cellpilot/${file}.bak"
    
    # Replace colors
    for old_color in "${!colors[@]}"; do
      new_color="${colors[$old_color]}"
      sed -i "s/${old_color}/${new_color}/gi" "/home/freddy/cellpilot/$file"
    done
    
    # Replace emojis
    for emoji in "${!emojis[@]}"; do
      replacement="${emojis[$emoji]}"
      sed -i "s/${emoji}/${replacement}/g" "/home/freddy/cellpilot/$file"
    done
    
    echo "  - Updated colors and removed emojis"
  fi
done

echo "Style update complete!"