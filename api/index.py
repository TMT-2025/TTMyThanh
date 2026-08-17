import sys
import os

# Append project root to sys.path so we can import app.py correctly
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
