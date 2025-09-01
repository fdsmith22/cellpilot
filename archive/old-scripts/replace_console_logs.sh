#!/bin/bash

# Replace console.log with Logger.info
find src/ -name "*.js" -type f -exec sed -i '' \
  -e 's/console\.log(/Logger.info(/g' \
  {} \;

# Replace console.error with Logger.error
find src/ -name "*.js" -type f -exec sed -i '' \
  -e 's/console\.error(/Logger.error(/g' \
  {} \;

# Replace console.warn with Logger.warn
find src/ -name "*.js" -type f -exec sed -i '' \
  -e 's/console\.warn(/Logger.warn(/g' \
  {} \;

# Replace console.debug with Logger.debug
find src/ -name "*.js" -type f -exec sed -i '' \
  -e 's/console\.debug(/Logger.debug(/g' \
  {} \;

echo "Console statements replaced with Logger calls"
echo "Files modified:"
grep -l "Logger\." src/*.js