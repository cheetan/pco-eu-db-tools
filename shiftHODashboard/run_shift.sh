#!/bin/bash
# Schedules the next wake then runs the dashboard refresh.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON=/opt/homebrew/bin/python3

bash "$SCRIPT_DIR/schedule_next_wake.sh"

caffeinate -i "$PYTHON" "$SCRIPT_DIR/dashboard.py" --refresh

# Upload XLSX files to BTP app
BASE=https://shift-ho-dashboard.cfapps.eu12-001.hana.ondemand.com
DIR=/Users/I324220/Library/CloudStorage/OneDrive-SAPSE/DesktopOneDrive/shiftHODashboard
EMAIL=catalin.popa@sap.com
BTP_UPLOAD_KEY=$(security find-generic-password -a "btp-upload" -s "shift-ho-dashboard" -w 2>/dev/null)

if [ -z "$BTP_UPLOAD_KEY" ]; then
  echo "[btp-upload] WARNING: upload key not found in keychain, skipping BTP upload"
else
  curl -sS -F "file=@$DIR/SnowAlert.xlsx"  -H "X-User-Email: $EMAIL" -H "X-Upload-Key: $BTP_UPLOAD_KEY" "$BASE/api/upload/snow_alert"  && echo " [snow_alert OK]"
  curl -sS -F "file=@$DIR/SnowCase.xlsx"   -H "X-User-Email: $EMAIL" -H "X-Upload-Key: $BTP_UPLOAD_KEY" "$BASE/api/upload/snow_case"   && echo " [snow_case OK]"
  curl -sS -F "file=@$DIR/SnowInc.xlsx"    -H "X-User-Email: $EMAIL" -H "X-Upload-Key: $BTP_UPLOAD_KEY" "$BASE/api/upload/snow_inc"    && echo " [snow_inc OK]"
  curl -sS -F "file=@$DIR/SPCReport.xlsx"  -H "X-User-Email: $EMAIL" -H "X-Upload-Key: $BTP_UPLOAD_KEY" "$BASE/api/upload/spc"         && echo " [spc OK]"
  curl -sS -F "file=@$DIR/MaxDB.xlsx"     -H "X-User-Email: $EMAIL" -H "X-Upload-Key: $BTP_UPLOAD_KEY" "$BASE/api/upload/spc_maxdb"   && echo " [spc_maxdb OK]"
fi
