#!/usr/bin/env python3
"""Create simple PNG icons for SheetMind Excel add-in."""

import base64
import os

# Simple 1x1 transparent PNG in base64
TRANSPARENT_PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg=="

def create_icon(filename):
    """Create a simple PNG icon file."""
    try:
        png_data = base64.b64decode(TRANSPARENT_PNG)
        with open(filename, 'wb') as f:
            f.write(png_data)
        print(f"✅ Created {filename}")
    except Exception as e:
        print(f"❌ Failed to create {filename}: {e}")

if __name__ == "__main__":
    # Create assets directory
    os.makedirs("assets", exist_ok=True)
    
    # Create icon files
    create_icon("assets/icon-16.png")
    create_icon("assets/icon-32.png") 
    create_icon("assets/icon-64.png")
    create_icon("assets/icon-80.png")
    
    print("🎨 Icon files created!") 