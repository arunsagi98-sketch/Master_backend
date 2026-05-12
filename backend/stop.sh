#!/bin/bash
# Stop the application
pkill -f "uvicorn backend.main:app"
pkill -f "gunicorn"

# Optional: Clean up temp files
rm -rf /tmp/*.xlsx