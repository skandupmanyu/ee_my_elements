#!/usr/bin/env python3
"""
Quick verification script to test if all dependencies are installed correctly.
"""

def verify_installation():
    """Verify that all required dependencies are installed."""
    print("🔍 Verifying installation...")
    
    try:
        import pptx
        print("✅ python-pptx library is installed")
        print(f"   Version: {pptx.__version__}")
    except ImportError:
        print("❌ python-pptx library is NOT installed")
        print("   Please run: pip install -r requirements.txt")
        return False
    
    try:
        import uuid
        print("✅ uuid module is available")
    except ImportError:
        print("❌ uuid module is NOT available")
        return False
    
    try:
        from pathlib import Path
        print("✅ pathlib module is available")
    except ImportError:
        print("❌ pathlib module is NOT available")
        return False
    
    print("\n🎉 All dependencies are installed correctly!")
    print("📝 You can now use the PowerPoint splitter:")
    print("   python pptx_slide_splitter.py your_presentation.pptx")
    
    return True

if __name__ == "__main__":
    verify_installation()
