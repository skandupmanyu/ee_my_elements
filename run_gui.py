#!/usr/bin/env python3
"""
PowerPoint Slide Splitter - GUI Launcher

Simple script to launch the Streamlit web GUI.

Author: AI Assistant
"""

import subprocess
import sys
import os
from pathlib import Path

def main():
    """Launch the Streamlit GUI."""
    
    # Get the directory of this script
    script_dir = Path(__file__).parent
    
    # Path to the Streamlit app
    streamlit_app = script_dir / "streamlit_app.py"
    
    if not streamlit_app.exists():
        print("❌ Error: streamlit_app.py not found!")
        sys.exit(1)
    
    print("🚀 Launching PowerPoint Slide Splitter GUI...")
    print("📱 The web interface will open in your default browser")
    print("🛑 Press Ctrl+C to stop the server")
    print()
    
    try:
        # Launch Streamlit
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", 
            str(streamlit_app),
            "--server.address", "localhost",
            "--server.port", "8501",
            "--browser.gatherUsageStats", "false"
        ])
    except KeyboardInterrupt:
        print("\n👋 GUI server stopped.")
    except Exception as e:
        print(f"❌ Error launching GUI: {e}")
        print("\n💡 Make sure Streamlit is installed:")
        print("   pip install streamlit")
        sys.exit(1)

if __name__ == "__main__":
    main()
