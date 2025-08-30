#!/bin/bash

# Launch script for Export for My Efficient Elements GUI
# This script activates the virtual environment and launches the Streamlit app

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$SCRIPT_DIR"

echo -e "${BLUE}🚀 Export for My Efficient Elements Launcher${NC}"
echo -e "${BLUE}================================================${NC}"

# Check if we're in the right directory
if [ ! -f "$PROJECT_DIR/scripts/run_gui.py" ]; then
    echo -e "${RED}❌ Error: Cannot find run_gui.py script${NC}"
    echo -e "${RED}   Make sure this script is in the project root directory${NC}"
    exit 1
fi

# Check if virtual environment exists
if [ ! -d "$PROJECT_DIR/venv" ]; then
    echo -e "${RED}❌ Error: Virtual environment not found${NC}"
    echo -e "${YELLOW}   Please run: python -m venv venv && source venv/bin/activate && pip install -r requirements.txt${NC}"
    exit 1
fi

# Change to project directory
cd "$PROJECT_DIR"

echo -e "${GREEN}📁 Project directory: $PROJECT_DIR${NC}"

# Activate virtual environment
echo -e "${GREEN}🔧 Activating virtual environment...${NC}"
source "$PROJECT_DIR/venv/bin/activate"

# Check if activation was successful
if [ -z "$VIRTUAL_ENV" ]; then
    echo -e "${RED}❌ Error: Failed to activate virtual environment${NC}"
    exit 1
fi

echo -e "${GREEN}✅ Virtual environment activated: $VIRTUAL_ENV${NC}"

# Check if required packages are installed
echo -e "${GREEN}🔍 Checking dependencies...${NC}"
python -c "import streamlit, pptx" 2>/dev/null
if [ $? -ne 0 ]; then
    echo -e "${YELLOW}⚠️  Installing missing dependencies...${NC}"
    pip install -r requirements.txt
fi

# Launch the GUI
echo -e "${GREEN}🚀 Launching GUI application...${NC}"
echo -e "${BLUE}📱 The web interface will open in your browser at http://localhost:8501${NC}"
echo -e "${BLUE}🛑 Press Ctrl+C in this terminal to stop the server${NC}"
echo ""

# Run the GUI script
python scripts/run_gui.py

# Cleanup message
echo ""
echo -e "${GREEN}👋 GUI application stopped.${NC}"
