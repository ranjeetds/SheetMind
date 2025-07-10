#!/usr/bin/env python3
"""
SheetMind Excel Add-in Setup Script

This script helps set up and serve the Excel add-in files.
"""

import http.server
import socketserver
import webbrowser
import os
import sys
from pathlib import Path


def setup_addin():
    """Set up and serve the Excel add-in."""
    print("🧠 SheetMind Excel Add-in Setup")
    print("=" * 40)
    
    # Check if we're in the right directory
    current_dir = Path.cwd()
    if not (current_dir / "manifest.xml").exists():
        print("❌ Error: manifest.xml not found!")
        print("Please run this script from the excel-addin directory.")
        sys.exit(1)
    
    # Configuration
    PORT = 3000
    BACKEND_PORT = 8000
    
    print(f"📁 Serving add-in files from: {current_dir}")
    print(f"🌐 Add-in will be available at: http://localhost:{PORT}")
    print(f"🔧 Backend should be running at: http://localhost:{BACKEND_PORT}")
    print()
    
    # Instructions
    print("📋 To install in Excel:")
    print("1. Open Excel")
    print("2. Go to Insert > Office Add-ins")
    print("3. Click 'Upload My Add-in'")
    print("4. Select manifest.xml from this folder")
    print("5. Click Upload")
    print()
    
    print("⚠️  Make sure your backend is running:")
    print(f"   python src/main.py web --port {BACKEND_PORT}")
    print()
    
    # Start server
    try:
        os.chdir(current_dir)
        
        class CustomHandler(http.server.SimpleHTTPRequestHandler):
            def do_GET(self):
                # Add CORS headers for Excel add-in
                self.send_response(200)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
                self.send_header('Access-Control-Allow-Headers', 'Content-Type')
                self.end_headers()
                return super().do_GET()
        
        with socketserver.TCPServer(("", PORT), CustomHandler) as httpd:
            print(f"🚀 Starting server on port {PORT}...")
            print("Press Ctrl+C to stop")
            print()
            
            # Open manifest location in file explorer (optional)
            try:
                manifest_path = current_dir / "manifest.xml"
                print(f"📄 Manifest file: {manifest_path}")
                print()
            except:
                pass
            
            httpd.serve_forever()
            
    except KeyboardInterrupt:
        print("\n👋 Server stopped. Goodbye!")
    except OSError as e:
        if e.errno == 48:  # Port already in use
            print(f"❌ Port {PORT} is already in use!")
            print("Try closing other applications or use a different port.")
        else:
            print(f"❌ Error starting server: {e}")
        sys.exit(1)


if __name__ == "__main__":
    setup_addin() 