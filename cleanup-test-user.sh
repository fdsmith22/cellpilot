#!/bin/bash

# Script to cleanup test user via API
EMAIL="$1"

if [ -z "$EMAIL" ]; then
  echo "Usage: ./cleanup-test-user.sh email@example.com"
  exit 1
fi

echo "Cleaning up user: $EMAIL"

# You need to be logged in as admin first
curl -X POST https://www.cellpilot.io/api/admin/cleanup-user \
  -H "Content-Type: application/json" \
  -d "{\"email\": \"$EMAIL\"}" \
  --cookie-jar cookies.txt \
  --cookie cookies.txt

echo "User cleaned up. They can now sign up again."