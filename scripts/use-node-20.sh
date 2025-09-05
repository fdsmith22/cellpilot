#!/bin/bash
# Script to run commands with Node v20

export PATH=/home/freddy/.nvm/versions/node/v20.19.5/bin:$PATH
export NODE_PATH=/home/freddy/.nvm/versions/node/v20.19.5/lib/node_modules

# Run the command passed as arguments
"$@"