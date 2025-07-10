#!/usr/bin/env python3
"""
Simple test script to verify SheetMind installation and basic functionality.

Run this script after installation to ensure everything is working correctly.
"""

import os
import sys
from pathlib import Path

def test_imports():
    """Test that all core modules can be imported."""
    print("🔍 Testing imports...")
    
    try:
        from src.agents.excel_agent import ExcelAgent
        from src.integrations.excel_handler import ExcelHandler
        from src.nlp.command_processor import CommandProcessor
        from src.ui.web.app import create_app
        print("✅ All core modules imported successfully")
        return True
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False

def test_dependencies():
    """Test that required dependencies are available."""
    print("\n🔍 Testing dependencies...")
    
    required_packages = [
        "fastapi", "uvicorn", "pandas", "openpyxl", 
        "rich", "click", "pydantic", "aiofiles"
    ]
    
    missing = []
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing.append(package)
    
    if missing:
        print(f"❌ Missing packages: {', '.join(missing)}")
        print("   Run: pip install -r requirements.txt")
        return False
    else:
        print("✅ All required dependencies available")
        return True

def test_agent_creation():
    """Test that ExcelAgent can be created."""
    print("\n🔍 Testing agent creation...")
    
    try:
        from src.agents.excel_agent import ExcelAgent
        agent = ExcelAgent()
        capabilities = agent.get_capabilities()
        
        if capabilities and len(capabilities) > 0:
            print(f"✅ Agent created successfully with {len(capabilities)} capabilities")
            return True
        else:
            print("❌ Agent created but no capabilities found")
            return False
    except Exception as e:
        print(f"❌ Agent creation failed: {e}")
        return False

def test_command_processing():
    """Test that command processing works."""
    print("\n🔍 Testing command processing...")
    
    try:
        from src.nlp.command_processor import CommandProcessor
        processor = CommandProcessor()
        
        # Test a simple command
        command = processor.process_command("calculate the sum of column A")
        
        if command.action == "calculate" and command.confidence > 0:
            print(f"✅ Command processed: {command.action} (confidence: {command.confidence:.2f})")
            return True
        else:
            print(f"❌ Command processing failed: {command.action}")
            return False
    except Exception as e:
        print(f"❌ Command processing error: {e}")
        return False

def test_web_app():
    """Test that web app can be created."""
    print("\n🔍 Testing web application...")
    
    try:
        from src.ui.web.app import create_app
        app = create_app()
        
        if app and hasattr(app, 'router'):
            print("✅ Web application created successfully")
            return True
        else:
            print("❌ Web application creation failed")
            return False
    except Exception as e:
        print(f"❌ Web app error: {e}")
        return False

def test_sample_data():
    """Test that sample data exists."""
    print("\n🔍 Testing sample data...")
    
    sample_file = Path("examples/sample_sales_data.csv")
    if sample_file.exists():
        print("✅ Sample data file found")
        return True
    else:
        print("❌ Sample data file not found")
        return False

def main():
    """Run all tests."""
    print("🧠 SheetMind Installation Test\n")
    print("=" * 50)
    
    tests = [
        test_dependencies,
        test_imports,
        test_agent_creation,
        test_command_processing,
        test_web_app,
        test_sample_data
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        if test():
            passed += 1
    
    print("\n" + "=" * 50)
    print(f"📊 Test Results: {passed}/{total} passed")
    
    if passed == total:
        print("\n🎉 All tests passed! SheetMind is ready to use.")
        print("\nNext steps:")
        print("1. Set up your API keys in .env file")
        print("2. Run: python src/main.py web")
        print("3. Open http://localhost:8000 in your browser")
        return True
    else:
        print(f"\n⚠️  {total - passed} tests failed. Please check the errors above.")
        print("\nTroubleshooting:")
        print("1. Make sure you've run: pip install -r requirements.txt")
        print("2. Check that you're using Python 3.8+")
        print("3. Verify all files are in the correct location")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1) 