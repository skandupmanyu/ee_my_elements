"""
Configuration settings for Export for My Efficient Elements.

This module contains all configurable settings for the application,
making it easy to modify behavior without changing core code.
"""

import os
from pathlib import Path

# Project paths
PROJECT_ROOT = Path(__file__).parent.parent
ASSETS_DIR = PROJECT_ROOT / "assets"
LOGO_PATH = ASSETS_DIR / "EfficientElementsLogo.png"

# Application settings
APP_NAME = "Export for My Efficient Elements"
APP_DESCRIPTION = "Convert your powerpoint deck into importable my elements"
DEFAULT_GROUP_NAME = "My Presentation"

# GUI settings
GUI_TITLE = APP_NAME
GUI_ICON = str(LOGO_PATH)
GUI_LAYOUT = "wide"  # "centered" or "wide"
GUI_SIDEBAR_STATE = "collapsed"  # "expanded" or "collapsed"

# Logo settings
LOGO_WIDTH = 150  # pixels
LOGO_CENTER = True

# File processing settings
DEFAULT_THUMBNAIL_HEIGHT = 300  # pixels for thumbnail generation
SUPPORTED_FILE_TYPES = ['pptx', 'ppt']
MAX_FILE_SIZE_MB = 200

# Output settings
XML_FILENAME = "MyElements.xml"
TEMP_DIR_PREFIX = "pptx_split_"
ZIP_TIMESTAMP_FORMAT = "%Y%m%d_%H%M%S"

# Progress settings
PROGRESS_UPDATE_DELAY = 0.2  # seconds between UI updates
ENABLE_VERBOSE_OUTPUT = True

# Thumbnail generation settings (PowerPoint to PDF to PNG pipeline)
THUMBNAIL_METHODS = {
    'powerpoint_applescript': {
        'name': 'Microsoft PowerPoint via AppleScript',
        'priority': 1,
        'timeout': 60
    },
    'keynote_applescript': {
        'name': 'Keynote via AppleScript',
        'priority': 2,
        'timeout': 60
    },
    'pdf2image': {
        'name': 'pdf2image library',
        'priority': 3,
        'timeout': 60
    },
    'poppler': {
        'name': 'Poppler utilities (pdftoppm)',
        'priority': 4,
        'timeout': 60
    },
    'simple_fallback': {
        'name': 'simple_fallback',
        'priority': 5,
        'timeout': None
    }
}

# Color scheme for UI
COLORS = {
    'primary': '#2E86C1',
    'secondary': '#7F8C8D',
    'success': '#27AE60',
    'warning': '#F39C12',
    'error': '#E74C3C',
    'info': '#3498DB',
    'dark': '#2C3E50',
    'purple': '#9B59B6',
    'brown': '#8B4513'
}

# Progress panel colors
PROGRESS_COLORS = {
    'creating_pptx': COLORS['info'],
    'creating_thumbnail': COLORS['warning'],
    'completed': COLORS['success'],
    'creating_xml': COLORS['purple'],
    'creating_zip': COLORS['brown'],
    'export_complete': COLORS['success']
}

# Environment-specific settings
DEBUG = os.getenv('DEBUG', 'False').lower() == 'true'
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')

# Performance settings
PARALLEL_PROCESSING = False  # Future feature
MAX_CONCURRENT_SLIDES = 3    # Future feature

def get_asset_path(filename: str) -> Path:
    """Get the full path to an asset file."""
    return ASSETS_DIR / filename

def get_logo_base64_path() -> str:
    """Get the logo path for base64 encoding."""
    return str(LOGO_PATH)

def get_app_config() -> dict:
    """Get application configuration as a dictionary."""
    return {
        'name': APP_NAME,
        'description': APP_DESCRIPTION,
        'logo_path': str(LOGO_PATH),
        'logo_width': LOGO_WIDTH,
        'supported_types': SUPPORTED_FILE_TYPES,
        'max_file_size_mb': MAX_FILE_SIZE_MB,
        'colors': COLORS,
        'debug': DEBUG
    }

def get_gui_config() -> dict:
    """Get GUI-specific configuration."""
    return {
        'title': GUI_TITLE,
        'icon': GUI_ICON,
        'layout': GUI_LAYOUT,
        'sidebar_state': GUI_SIDEBAR_STATE,
        'logo_width': LOGO_WIDTH,
        'colors': COLORS,
        'progress_colors': PROGRESS_COLORS
    }

def get_processing_config() -> dict:
    """Get processing-specific configuration."""
    return {
        'thumbnail_height': DEFAULT_THUMBNAIL_HEIGHT,
        'xml_filename': XML_FILENAME,
        'temp_prefix': TEMP_DIR_PREFIX,
        'timestamp_format': ZIP_TIMESTAMP_FORMAT,
        'thumbnail_methods': THUMBNAIL_METHODS,
        'progress_delay': PROGRESS_UPDATE_DELAY,
        'verbose': ENABLE_VERBOSE_OUTPUT
    }
