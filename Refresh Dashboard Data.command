#!/bin/zsh
set -e

cd "/Users/vbond/dashboard"

if [ -d "/Users/vbond/.vendor" ]; then
  export PYTHONPATH="/Users/vbond/.vendor"
fi

python3 scripts/build_dashboard_data.py

echo ""
echo "Dashboard data.json has been refreshed."
echo "Press Enter to close."
read
