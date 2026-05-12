#!/bin/bash
# Production startup script for Hostinger VPS

# Activate virtual environment
source venv/bin/activate

# Set environment variables
export PYTHONPATH=/home/u123456789/mpn-generator
export HOST=0.0.0.0
export PORT=8000

# Start the application with Gunicorn
gunicorn backend.main:app -w 4 -k uvicorn.workers.UvicornWorker --bind 0.0.0.0:8000 --daemon

# Alternative: Run with uvicorn directly
# uvicorn backend.main:app --host 0.0.0.0 --port 8000 --workers 4 --daemon