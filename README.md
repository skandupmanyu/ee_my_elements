# PowerPoint Slide Splitter

A Python application that takes a PowerPoint file as input and converts each slide into individual .pptx files, each named with a unique UUID.

## Features

- ✅ Splits PowerPoint presentations (.pptx, .ppt) into individual slide files
- ✅ Each output file is named with a unique UUID
- ✅ Preserves slide content and formatting
- ✅ **NEW:** Generates high-quality PNG thumbnail images (120px height)
- ✅ Generates XML metadata file (`MyElements.xml`) with slide information
- ✅ **NEW:** Creates clean timestamped zip archive and auto-cleanup
- ✅ **NEW:** User-friendly web GUI with drag-and-drop file upload
- ✅ Reproducible group UUIDs based on group names
- ✅ Automatic slide title extraction for meaningful element names
- ✅ Both web GUI and command-line interfaces available
- ✅ Customizable output directory and group names

## Installation

1. Clone or download this repository
2. Create a virtual environment (recommended):

```bash
python -m venv venv
```

3. Activate the virtual environment:

**On macOS/Linux:**
```bash
source venv/bin/activate
```

**On Windows:**
```bash
venv\Scripts\activate
```

4. Install the required dependencies:

```bash
pip install -r requirements.txt
```

5. To deactivate the virtual environment when done:

```bash
deactivate
```

6. (Optional) Verify the installation:

```bash
python verify_install.py
```

## Usage

You can use this tool in two ways: through a user-friendly web GUI or via command line for automation and debugging.

### 🖥️ Web GUI (Recommended for Most Users)

**Quick Start:**
```bash
# First, activate your virtual environment
source venv/bin/activate  # On macOS/Linux
# or
venv\Scripts\activate     # On Windows

# Launch the web GUI
python run_gui.py
```

This will launch a web interface in your browser where you can:
- 📁 **Drag & drop** your PowerPoint file
- 🏷️ **Enter a group name** for your slides
- 🚀 **Click process** and get your zip file automatically
- 📥 **Download** the result directly

**GUI Features:**
- ✅ **User-friendly interface** - No command-line knowledge needed
- ✅ **Real-time progress** - See processing status as it happens
- ✅ **Automatic downloads** - Get your zip file instantly
- ✅ **Error handling** - Clear error messages if something goes wrong
- ✅ **File validation** - Ensures your file is a valid PowerPoint presentation
- ✅ **Mobile-friendly** - Works on tablets and mobile devices
- ✅ **Local processing** - Files processed on your machine, not uploaded to servers

### ⌨️ Command Line (For Automation & Debugging)

**Basic Usage:**

First, activate your virtual environment:
```bash
source venv/bin/activate  # On macOS/Linux
# or
venv\Scripts\activate     # On Windows
```

Then run the splitter:
```bash
python pptx_slide_splitter.py presentation.pptx
```

This will process the slides using a temporary directory and create a clean zip archive next to your input file.

### Specify Output Directory

```bash
python pptx_slide_splitter.py presentation.pptx -o my_slides/
```

### Specify Group Name for XML Metadata

```bash
python pptx_slide_splitter.py presentation.pptx -g "PodHandler"
```

### Combined Options

```bash
python pptx_slide_splitter.py presentation.pptx -g "My Custom Group" -o my_slides/ -v
```

### Help

```bash
python pptx_slide_splitter.py --help
```

## Examples

### Split a presentation into individual slides with custom group:
```bash
python pptx_slide_splitter.py my_presentation.pptx -g "PodHandler" -v
```

Output:
```
Loading presentation: my_presentation.pptx
Found 5 slides to process
Processing slide 1/5
  → Created thumbnail: a1b2c3d4-e5f6-7890-abcd-ef1234567890.png
  → Saved as: a1b2c3d4-e5f6-7890-abcd-ef1234567890.pptx
Processing slide 2/5
  → Created thumbnail: b2c3d4e5-f6g7-8901-bcde-f23456789012.png
  → Saved as: b2c3d4e5-f6g7-8901-bcde-f23456789012.pptx
...
📄 Created XML metadata: MyElements.xml
   Group: PodHandler (ID: 7bbf14b3-a82a-568c-a3ef-ab5ad0171dbe)
   Elements: 5

📦 Creating zip archive...
    ✅ Compressed 11 files
    📦 Archive size: 12.3 MB
    📁 Saved to: /path/to/presentation_20240829_143022.zip

🧹 Cleaning up generated files...
    ✅ Removed 11 generated files
    🗑️  Removed temporary directory: pptx_split_abc123
    📦 Final output: presentation_20240829_143022.zip

🎉 Processing complete!
⏱️  Total time: 15.6s (3.1s per slide)
🚀 Performance: 0.3 slides/second
📈 High-quality thumbnails with accurate visual representation!
```

### Specify a custom output directory:
```bash
python pptx_slide_splitter.py my_presentation.pptx --output-dir ./individual_slides/
```

## Output

### Individual Slide Files
Each slide will be saved as a separate .pptx file with a UUID filename format:
- `a1b2c3d4-e5f6-7890-abcd-ef1234567890.pptx`
- `b2c3d4e5-f6g7-8901-bcde-f23456789012.pptx`
- etc.

The UUID ensures each file has a unique name, preventing any conflicts.

### Thumbnail Images (NEW!)
For each slide, a high-quality PNG thumbnail is generated with the same UUID:
- `a1b2c3d4-e5f6-7890-abcd-ef1234567890.png`
- `b2c3d4e5-f6g7-8901-bcde-f23456789012.png`
- etc.

**Thumbnail Features:**
- **120px height** with maintained aspect ratio
- **Pixel-perfect rendering** of slide content
- **All visual elements preserved** (colors, layouts, shapes, styling)
- **High quality PNG format** with 95% quality

### Final Output: Clean Zip Archive (NEW!)
**Location:** Placed next to your original PowerPoint file  
**Filename:** `your_presentation_YYYYMMDD_HHMMSS.zip`

**Example:**
```
/Users/you/Documents/
├── my_presentation.pptx          ← Your original file
└── my_presentation_20240829_143022.zip  ← Generated archive
```

**Archive Contents (Root Level):**
- **All individual PPTX files** (slide files)
- **All PNG thumbnails** (one per slide)
- **XML metadata file** (`MyElements.xml`)
- **No system files** (excludes .DS_Store, Thumbs.db, etc.)
- **Root-level structure** (no nested folders inside zip)

**Automatic Cleanup:**
- ✅ **Temporary files removed** - No clutter left behind
- ✅ **Single output file** - Just one clean zip archive
- ✅ **Optimal location** - Zip placed next to original file

**Benefits:**
- **🎯 Single file output** - Easy to share and manage
- **🧹 Clean workspace** - No scattered temporary files
- **📦 Professional packaging** - All content organized in one archive
- **⚡ Fast access** - Zip placed conveniently next to source file
- **💾 Efficient storage** - Compressed for optimal file size

### XML Metadata File
A `MyElements.xml` file is automatically generated with the following structure:

```xml
<ee4p>
  <group id="7bbf14b3-a82a-568c-a3ef-ab5ad0171dbe" name="PodHandler">
    <element name="Unlocking Growth and Loyalty with Personalizati..." thumbMode="1" id="bc893a1b-c4f5-428e-8b2c-a34a5008f3b8"/>
    <element name="Customer acquisition" thumbMode="1" id="f41bf0a5-9019-4417-ba21-0e6693fda57b"/>
    <element name="Fabriq deployed to scale personalization and me..." thumbMode="1" id="6f6e2ed1-aa4c-4605-89bc-414fbdb55725"/>
    ...
  </group>
</ee4p>
```

**XML Structure:**
- **No XML declaration:** File starts directly with the root element (no `<?xml version="1.0" ?>`)
- **Root element:** `<ee4p>`
- **Group element:** Contains all slides from the presentation
  - `id`: Reproducible UUID based on group name (same group name = same UUID)
  - `name`: The group name you specify (or presentation filename if not specified)
- **Element elements:** One for each split slide
  - `name`: Extracted slide title or "Slide X" if no title found
  - `thumbMode`: Always set to "1"
  - `id`: Same UUID as the corresponding .pptx filename

## Supported File Formats

- Input: `.pptx` and `.ppt` files
- Output: `.pptx` files

## Requirements

- Python 3.6+
- python-pptx library
- Pillow (PIL) library for image processing

## Error Handling

The application includes comprehensive error handling for:
- Missing input files
- Invalid file formats
- Permission issues
- Corrupted PowerPoint files
- Thumbnail generation failures (uses fallback thumbnails)

## Technical Details

The application uses efficient libraries for fast processing:

**PowerPoint Processing (`python-pptx`):**
1. Load the source PowerPoint presentation
2. Iterate through each slide
3. Create a new presentation for each slide
4. Copy the slide content while preserving formatting
5. Save each new presentation with a UUID filename

**High-Quality Thumbnail Generation:**
1. **macOS Quick Look** for pixel-perfect slide rendering (preferred method)
2. **LibreOffice fallback** for accurate cross-platform conversion
3. **Smart method detection** automatically chooses best available option
4. **Single thumbnail** at 120px height with maintained aspect ratio
5. **High-quality PNG output** with same UUID as slide

**XML Metadata:**
1. Extract slide titles using text analysis
2. Generate reproducible UUIDs for groups using UUID5
3. Create structured XML with slide information

**Zip Archive Creation & Cleanup:**
1. **Temporary directory** - uses system temp directory for intermediate files
2. **Automatic compression** of all generated files
3. **Timestamp-based naming** using input filename + YYYYMMDD_HHMMSS format
4. **Strategic placement** - zip saved next to original PowerPoint file
5. **Selective inclusion** - only PPTX, PNG, and XML files
6. **Root-level structure** - files added directly to zip root
7. **Optimal compression** using ZIP_DEFLATED with compression level 6
8. **Automatic cleanup** - removes all temporary files and directories after archiving
9. **Clean workspace** - no intermediate files left behind

### Quality Benefits:
- **Pixel-perfect thumbnails** that accurately represent slide content
- **Automatic method selection** uses best available conversion tool
- **Preserves visual elements** including colors, layouts, shapes, and styling
- **Cross-platform compatibility** with graceful fallbacks
- **Professional output** suitable for production use

## License

This project is open source and available under the MIT License.

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.
