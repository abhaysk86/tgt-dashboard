#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# TGT AOA Dashboard Updater
# Run this any time you want to refresh the dashboard.
# ─────────────────────────────────────────────────────────────────────────────

set -e
cd "$(dirname "$0")"

echo "Generating dashboard..."
python3 generate.py

echo "Pushing to GitHub..."
git add index.html
git commit -m "Dashboard update $(date '+%d %b %Y')"
git push

echo ""
echo "Done! Your dashboard is live."
