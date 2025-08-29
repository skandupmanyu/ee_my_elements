# Export for My Efficient Elements

A professional Python application that converts PowerPoint presentations into individual slide files with high-quality thumbnails and XML metadata for seamless integration with presentation software.

![Logo](assets/EfficientElementsLogo.png)

## ✨ Features

- **🎯 Dual Interface**: Both command-line and web GUI interfaces
- **📄 Individual Slides**: Each slide becomes a separate PPTX file with unique UUID naming
- **🖼️ High-Quality Thumbnails**: Professional PNG previews using macOS Quick Look
- **📋 XML Metadata**: Structured MyElements.xml for easy importing
- **📦 Clean Packaging**: Everything bundled in timestamped ZIP archives
- **⚡ Real-Time Progress**: Live progress tracking with detailed status updates
- **🧹 Automatic Cleanup**: Clean intermediate file handling
- **⚙️ Configurable**: Centralized configuration system for easy customization

## 🏗️ Project Structure

```
ee_my_elements/
├── README.md                    # Project documentation
├── requirements.txt             # Python dependencies
├── .gitignore                   # Git ignore rules
├── config/                      # Configuration management
│   ├── __init__.py
│   └── settings.py              # Centralized application settings
├── src/                         # Source code
│   ├── __init__.py
│   ├── core/                    # Core business logic
│   │   ├── __init__.py
│   │   ├── splitter.py          # Main PowerPointSplitter class
│   │   ├── thumbnail_generator.py  # High-quality thumbnail generation
│   │   └── xml_generator.py     # XML metadata creation
│   ├── gui/                     # Web interface
│   │   ├── __init__.py
│   │   └── streamlit_app.py     # Streamlit web application
│   └── utils/                   # Utility functions
│       ├── __init__.py
│       ├── file_utils.py        # File operations and archive creation
│       └── uuid_utils.py        # UUID generation utilities
├── assets/                      # Static assets
│   └── EfficientElementsLogo.png  # Application logo
├── scripts/                     # Entry point scripts
│   ├── run_cli.py              # Command-line interface launcher
│   ├── run_gui.py              # Web GUI launcher
│   └── verify_install.py       # Installation verification
└── tests/                       # Test directory (future)
    └── __init__.py
```

## 🚀 Quick Start

### Prerequisites

- **Python 3.8+** (Python 3.12+ recommended)
- **macOS** (required for thumbnail generation via Quick Look)

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/skandupmanyu/ee_my_elements.git
   cd ee_my_elements
   ```

2. **Create and activate virtual environment:**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On macOS/Linux
   # or
   venv\Scripts\activate     # On Windows
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Verify installation:**
   ```bash
   python scripts/verify_install.py
   ```

## 💻 Usage

### Web Interface (Recommended)

Launch the user-friendly web interface:

```bash
python scripts/run_gui.py
```

Then open `http://localhost:8501` in your browser to:
- Upload PowerPoint files via drag-and-drop
- Set folder names for organization
- Watch real-time processing progress
- Download the final zip archive

### Command Line Interface

For automation and advanced usage:

```bash
# Basic usage
python scripts/run_cli.py presentation.pptx

# With custom settings
python scripts/run_cli.py presentation.pptx \
  --group-name "My Project" \
  --output-dir ./output \
  --verbose

# See all options
python scripts/run_cli.py --help
```

### CLI Options

- `input_file`: Path to PowerPoint file (.pptx or .ppt)
- `-g, --group-name`: Name for XML metadata grouping
- `-o, --output-dir`: Custom output directory (optional)
- `-b, --base-name`: Custom base name for zip file
- `-v, --verbose`: Enable detailed output
- `--debug`: Show detailed error information

## ⚙️ Configuration

The application uses a centralized configuration system in `config/settings.py`. Key settings include:

### Application Settings
```python
APP_NAME = "Export for My Efficient Elements"
DEFAULT_GROUP_NAME = "My Presentation"
SUPPORTED_FILE_TYPES = ['pptx', 'ppt']
MAX_FILE_SIZE_MB = 200
```

### Thumbnail Settings
```python
DEFAULT_THUMBNAIL_HEIGHT = 120
THUMBNAIL_METHODS = {
    'macos_quicklook': {'priority': 1, 'timeout': 30},
    'libreoffice': {'priority': 2, 'timeout': 60},
    'pil_fallback': {'priority': 3}
}
```

### UI Colors
```python
COLORS = {
    'primary': '#2E86C1',
    'success': '#27AE60',
    'warning': '#F39C12',
    'error': '#E74C3C'
}
```

## 📋 Output Structure

Each processing run creates:

### Individual Files
- **`{uuid}.pptx`**: Individual slide presentations
- **`{uuid}.png`**: High-quality thumbnails (120px height)
- **`MyElements.xml`**: Metadata for importing

### Final Archive
- **`{filename}_{timestamp}.zip`**: Complete package ready for import

### XML Structure
```xml
<ee4p>
  <group id="{reproducible-uuid}" name="{group-name}">
    <element name="{slide-title}" thumbMode="1" id="{slide-uuid}"/>
    <!-- More elements... -->
  </group>
</ee4p>
```

## 🔧 Development

### Adding New Features

1. **Core Logic**: Add to `src/core/`
2. **Utilities**: Add to `src/utils/`
3. **Configuration**: Update `config/settings.py`
4. **GUI Components**: Modify `src/gui/streamlit_app.py`

### Testing

```bash
# Test all imports
python -c "from src.core.splitter import PowerPointSplitter; print('✅ All imports working')"

# Test CLI
python scripts/run_cli.py --version

# Test GUI
python scripts/run_gui.py
```

### Code Organization

- **`config/`**: All configurable settings
- **`src/core/`**: Business logic (splitting, thumbnails, XML)
- **`src/utils/`**: Reusable utilities (files, UUIDs)
- **`src/gui/`**: User interface components
- **`scripts/`**: Entry points and launchers

## 🎯 Integration with PowerPoint

To import the generated elements:

1. **Extract** the downloaded zip file
2. **Open PowerPoint**
3. **Click** on Bugs or Icons button to open element wizard
4. **Navigate** to "My elements" in the bottom of left panel
5. **Click** import button at the bottom
6. **Select** the downloaded zip file

## 🛠️ Troubleshooting

### Common Issues

**Import Errors:**
```bash
# Ensure you're in the project root and virtual environment is active
source venv/bin/activate
export PYTHONPATH=$PWD:$PYTHONPATH
```

**Thumbnail Quality Issues:**
- Ensure you're running on macOS for optimal Quick Look integration
- Check that `qlmanage` command is available in your system

**File Size Limits:**
- Default limit: 200MB (configurable in `config/settings.py`)
- Large files may require more processing time

### Performance Tips

- **macOS Quick Look**: Provides fastest, highest-quality thumbnails with no additional setup
- **Optimized for macOS**: Streamlined codebase with minimal dependencies
- **Large files**: Enable verbose mode to monitor progress

## 📝 Dependencies

### Core Dependencies
- **python-pptx**: PowerPoint file manipulation (includes Pillow dependency)
- **streamlit**: Web interface framework

### System Requirements
- **macOS Quick Look**: Built-in thumbnail generation (no additional installation required)

**Note**: Pillow is included as a dependency because python-pptx requires it internally, but our thumbnail generation is streamlined to use only macOS Quick Look for optimal quality and performance.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Follow the existing code structure
4. Update configuration in `config/settings.py` as needed
5. Test both CLI and GUI interfaces
6. Submit a pull request

## 📄 License

This project is open source. See the repository for license details.

## 🔗 Links

- **GitHub Repository**: https://github.com/skandupmanyu/ee_my_elements
- **Issues & Support**: Use GitHub Issues for bug reports and feature requests

---

**Built with ❤️ for seamless PowerPoint element management**
