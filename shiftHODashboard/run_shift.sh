#!/bin/bash
# Schedules the next wake then runs the dashboard refresh.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON=/opt/homebrew/bin/python3

bash "$SCRIPT_DIR/schedule_next_wake.sh"

caffeinate -i "$PYTHON" "$SCRIPT_DIR/dashboard.py" --refresh

# Upload XLSX files to BTP app
BASE=https://shift-ho-dashboard.cfapps.eu01-canary.hana.ondemand.com
DIR=/Users/I324220/Library/CloudStorage/OneDrive-SAPSE/DesktopOneDrive/shiftHODashboard
EMAIL=catalin.popa@sap.com
BTP_UPLOAD_KEY=$(security find-generic-password -a "btp-upload" -s "shift-ho-dashboard" -w 2>/dev/null)

if [ -z "$BTP_UPLOAD_KEY" ]; then
  echo "[btp-upload] WARNING: upload key not found in keychain, skipping BTP upload"
else
  echo "[btp-upload] Waiting 10s for OneDrive to finish syncing..."
  sleep 10
  for src_file in SnowAlert:snow_alert SnowCase:snow_case SnowInc:snow_inc SPCReport:spc MaxDB:spc_maxdb; do
    FNAME="${src_file%%:*}.xlsx"
    SOURCE="${src_file##*:}"
    if [ ! -f "$DIR/$FNAME" ]; then
      echo "[btp-upload] WARNING: $FNAME not found, skipping $SOURCE"
      continue
    fi
    HTTP_CODE=$(curl -s -o /dev/null -w "%{http_code}" \
      -F "file=@$DIR/$FNAME" \
      -H "X-User-Email: $EMAIL" \
      -H "X-Upload-Key: $BTP_UPLOAD_KEY" \
      "$BASE/api/upload/$SOURCE")
    if [ "$HTTP_CODE" = "200" ]; then
      echo " [$SOURCE OK]"
    else
      echo " [$SOURCE FAILED HTTP $HTTP_CODE]"
    fi
  done
fi
