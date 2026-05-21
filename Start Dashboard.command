#!/bin/zsh
set -e

cd "/Users/vbond/dashboard"

if [ -d "/Users/vbond/.vendor" ]; then
  export PYTHONPATH="/Users/vbond/.vendor"
fi

if ! pgrep -f "scripts/dashboard_server.py" >/dev/null 2>&1; then
  nohup python3 scripts/dashboard_server.py >/tmp/lg_dashboard_server.log 2>&1 &
  sleep 2
fi

open "http://127.0.0.1:8000/"
