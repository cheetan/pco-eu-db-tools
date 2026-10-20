#!/bin/bash
# Calculates the next shift wake time (07:50, 15:50, 23:50) and schedules it with pmset.
PYTHON=/opt/homebrew/bin/python3

NEXT=$("$PYTHON" - <<'EOF'
from datetime import datetime, timedelta
now = datetime.now()
times = [(7, 48), (15, 48), (23, 48)]
candidates = []
for h, m in times:
    t = now.replace(hour=h, minute=m, second=0, microsecond=0)
    if t <= now:
        t += timedelta(days=1)
    candidates.append(t)
print(min(candidates).strftime('%m/%d/%Y %H:%M:%S'))
EOF
)

sudo /usr/bin/pmset schedule wake "$NEXT"
echo "[$(date)] Wake scheduled for: $NEXT"
