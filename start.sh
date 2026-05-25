#!/bin/bash
echo "🚀 مذكرتي Pro v18 — Starting..."

# Init DB
python3 -c "from core.database import init_db; init_db(); print('✅ Database ready')"

# Start server
exec gunicorn app:app \
  --bind 0.0.0.0:${PORT:-5000} \
  --workers 2 \
  --timeout 120 \
  --preload \
  --access-logfile - \
  --error-logfile -
