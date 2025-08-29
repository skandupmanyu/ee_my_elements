#!/usr/bin/env python3
"""
GUI launcher for Export for My Efficient Elements.

This script launches the Streamlit web interface for user-friendly
PowerPoint processing with real-time progress tracking.
"""

import subprocess
import sys
import os
from pathlib import Path

def main():
    """Launch the Streamlit GUI application."""
    
    # Get the project root directory (parent of scripts directory)
    project_root = Path(__file__).parent.parent
    
    # Path to the Streamlit app in the new structure
    streamlit_app = project_root / "src" / "gui" / "streamlit_app.py"
    
    if not streamlit_app.exists():
        print(f"❌ Error: Streamlit app not found at {streamlit_app}")
        print("💡 Make sure the project structure is correct")
        sys.exit(1)
    
    print("🚀 Launching Export for My Efficient Elements GUI...")
    print("📱 The web interface will open in your default browser")
    print("🛑 Press Ctrl+C to stop the server")
    print()
    
    # Change to project directory for proper imports
    os.chdir(project_root)
    
    # Add project root to Python path for imports
    env = os.environ.copy()
    if 'PYTHONPATH' in env:
        env['PYTHONPATH'] = f"{project_root}:{env['PYTHONPATH']}"
    else:
        env['PYTHONPATH'] = str(project_root)
    
    try:
        # Launch Streamlit with the new app location
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", 
            str(streamlit_app),
            "--server.address", "localhost",
            "--server.port", "8501",
            "--browser.gatherUsageStats", "false"
        ], env=env)
        
    except KeyboardInterrupt:
        print("\n👋 GUI server stopped.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error launching GUI: {e}")
        print("💡 Make sure Streamlit is installed: pip install streamlit")
        print("💡 Check that all dependencies are available")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
