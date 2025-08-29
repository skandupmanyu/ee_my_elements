# PowerPoint Slide Splitter

A Python application that takes a PowerPoint file as input and converts each slide into individual .pptx files, each named with a unique UUID.

## Features

- ✅ Splits PowerPoint presentations (.pptx, .ppt) into individual slide files
- ✅ Each output file is named with a unique UUID
- ✅ Preserves slide content and formatting
- ✅ **NEW:** Generates high-quality PNG thumbnail images (120px height)
- ✅ Generates XML metadata file (`MyElements.xml`) with slide information
- ✅ Reproducible group UUIDs based on group names
- ✅ Automatic slide title extraction for meaningful element names
- ✅ Command-line interface for easy usage
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

### Basic Usage

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

This will create individual slide files in a `split_slides/` directory in the same location as your input file.

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

✅ Successfully created 5 individual slide files
📁 Output directory: split_slides/

Created files:
  • a1b2c3d4-e5f6-7890-abcd-ef1234567890.pptx
  • a1b2c3d4-e5f6-7890-abcd-ef1234567890.png
  • b2c3d4e5-f6g7-8901-bcde-f23456789012.pptx
  • b2c3d4e5-f6g7-8901-bcde-f23456789012.png
  • MyElements.xml
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
