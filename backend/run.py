#!/usr/bin/env python3
"""
Main entry point for the MPN Campaign Report Generator application
"""
import sys
from pathlib import Path

# Add backend directory to Python path
backend_dir = Path(__file__).parent / "backend"
sys.path.insert(0, str(backend_dir.parent))

# Import and run the app
from backend.main import app
import uvicorn

if __name__ == "__main__":
    print("Starting MPN Campaign Report Generator...")
    print("Open http://localhost:8000 in your browser")
    uvicorn.run(app, host="0.0.0.0", port=8000)
