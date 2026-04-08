#!/bin/bash
set -e

DB_PATH="/app/data/packinglist.db"
mkdir -p /app/data

# Restore the database from Azure Blob Storage if a backup exists.
# -if-db-not-exists: skip restore if DB already present (e.g. container reuse).
# -if-replica-exists: skip if no backup exists yet (first run).
echo "Restoring database from LiteStream replica..."
litestream restore -if-db-not-exists -if-replica-exists -config /app/litestream.yml "$DB_PATH"

# Start uvicorn under litestream replicate so writes are streamed continuously.
echo "Starting app with LiteStream replication..."
exec litestream replicate -config /app/litestream.yml -exec "uvicorn main:app --host 0.0.0.0 --port 8000"
