"""
Thumbnail generation module for Export for My Efficient Elements.

This module handles thumbnail generation using a reliable two-step approach:
1. PowerPoint to PDF conversion (using Microsoft PowerPoint via AppleScript)
2. PDF to PNG conversion (using pdf2image library)

This approach provides high-quality, accurate slide thumbnails.
"""

import os
import subprocess
import tempfile
import time
from pathlib import Path
from typing import List, Optional

from config.settings import get_processing_config


class SlideThumbnailGenerator:
    """High-quality thumbnail generator using PowerPoint to PDF to PNG conversion."""
    
    def __init__(self):
        self.config = get_processing_config()
        
        # Check available conversion methods
        self.conversion_methods = self._detect_conversion_methods()
        if self.config['verbose']:
            print(f"🔍 Available conversion methods: {', '.join(self.conversion_methods)}")
    
    def _detect_conversion_methods(self) -> List[str]:
        """Detect which conversion methods are available on this system."""
        methods = []
        
        # Check for Microsoft PowerPoint (AppleScript method)
        try:
            result = subprocess.run(['osascript', '-e', 'tell application "Microsoft PowerPoint" to get version'], 
                                  capture_output=True, timeout=5)
            if result.returncode == 0:
                methods.append('powerpoint_applescript')
        except:
            pass
        
        # Check for Keynote (AppleScript method)
        try:
            result = subprocess.run(['osascript', '-e', 'tell application "Keynote" to get version'], 
                                  capture_output=True, timeout=5)
            if result.returncode == 0:
                methods.append('keynote_applescript')
        except:
            pass
        
        # Check for pdf2image library
        try:
            from pdf2image import convert_from_path
            methods.append('pdf2image')
        except ImportError:
            pass
        
        # Check for Poppler (pdftoppm)
        try:
            result = subprocess.run(['pdftoppm', '-h'], capture_output=True, timeout=3)
            if result.returncode == 0:
                methods.append('poppler')
        except:
            pass
        
        # Always available simple fallback
        methods.append('simple_fallback')
        
        return methods
    
    def create_high_quality_thumbnail_from_pptx(self, pptx_path: str, slide_number: int) -> Optional[str]:
        """
        Create a high-quality thumbnail from PPTX file using PowerPoint to PDF to PNG conversion.
        
        Args:
            pptx_path: Path to the PPTX file
            slide_number: Slide number for identification
            
        Returns:
            Path to the generated thumbnail file, or None if failed
        """
        if self.config['verbose']:
            print(f"    🎨 Generating high-quality thumbnail using PowerPoint to PNG conversion...")
        
        # Try PowerPoint to PDF to PNG conversion
        thumbnail_path = self._convert_ppt_to_png(pptx_path, slide_number)
        if thumbnail_path:
            if self.config['verbose']:
                print(f"    ✅ Used PowerPoint to PNG conversion for high-quality thumbnail")
            return thumbnail_path
        
        # Fallback to simple placeholder
        if self.config['verbose']:
            print(f"    ℹ️  Using simple fallback thumbnail")
        return self._create_simple_fallback_thumbnail(pptx_path, slide_number)
    
    def _convert_ppt_to_png(self, pptx_path: str, slide_number: int) -> Optional[str]:
        """Convert PPTX to PNG using PowerPoint to PDF to PNG pipeline."""
        
        # Step 1: Convert PPT to PDF
        pdf_path = self._convert_ppt_to_pdf(pptx_path)
        if not pdf_path:
            return None
        
        try:
            # Step 2: Convert PDF to PNG
            png_path = self._convert_pdf_to_png(pdf_path, slide_number)
            return png_path
        finally:
            # Clean up temporary PDF file
            try:
                if pdf_path and Path(pdf_path).exists():
                    Path(pdf_path).unlink()
            except:
                pass
    
    def _convert_ppt_to_pdf(self, pptx_path: str) -> Optional[str]:
        """Convert PowerPoint to PDF using the best available method."""
        
        # Try Microsoft PowerPoint via AppleScript first
        if 'powerpoint_applescript' in self.conversion_methods:
            pdf_path = self._convert_ppt_to_pdf_applescript_powerpoint(pptx_path)
            if pdf_path:
                return pdf_path
        
        # Try Keynote via AppleScript as fallback
        if 'keynote_applescript' in self.conversion_methods:
            pdf_path = self._convert_ppt_to_pdf_applescript_keynote(pptx_path)
            if pdf_path:
                return pdf_path
        
        return None
    
    def _convert_ppt_to_pdf_applescript_powerpoint(self, pptx_path: str) -> Optional[str]:
        """Convert PowerPoint to PDF using Microsoft PowerPoint via AppleScript."""
        
        try:
            # Create temporary PDF file
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
                pdf_path = temp_file.name
            
            # AppleScript to convert PPT to PDF using PowerPoint
            applescript = f'''
            tell application "Microsoft PowerPoint"
                open POSIX file "{pptx_path}"
                set thePresentation to active presentation
                save thePresentation in POSIX file "{pdf_path}" as save as PDF
                close thePresentation
            end tell
            '''
            
            # Run AppleScript
            result = subprocess.run(
                ["osascript", "-e", applescript],
                capture_output=True,
                text=True,
                timeout=60
            )
            
            if result.returncode == 0 and Path(pdf_path).exists():
                return pdf_path
            else:
                # Clean up failed attempt
                try:
                    Path(pdf_path).unlink()
                except:
                    pass
                return None
                
        except Exception:
            return None
    
    def _convert_ppt_to_pdf_applescript_keynote(self, pptx_path: str) -> Optional[str]:
        """Convert PowerPoint to PDF using Keynote via AppleScript."""
        
        try:
            # Create temporary PDF file
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
                pdf_path = temp_file.name
            
            # AppleScript to convert PPT to PDF using Keynote
            applescript = f'''
            tell application "Keynote"
                open POSIX file "{pptx_path}"
                set thePresentation to front document
                export thePresentation to file "{pdf_path}" as PDF
                close thePresentation
            end tell
            '''
            
            # Run AppleScript
            result = subprocess.run(
                ["osascript", "-e", applescript],
                capture_output=True,
                text=True,
                timeout=60
            )
            
            if result.returncode == 0 and Path(pdf_path).exists():
                return pdf_path
            else:
                # Clean up failed attempt
                try:
                    Path(pdf_path).unlink()
                except:
                    pass
                return None
                
        except Exception:
            return None
    
    def _convert_pdf_to_png(self, pdf_path: str, slide_number: int) -> Optional[str]:
        """Convert PDF to PNG using the best available method."""
        
        # Try pdf2image first (most reliable)
        if 'pdf2image' in self.conversion_methods:
            png_path = self._convert_pdf_to_png_pdf2image(pdf_path, slide_number)
            if png_path:
                return png_path
        
        # Try Poppler as fallback
        if 'poppler' in self.conversion_methods:
            png_path = self._convert_pdf_to_png_poppler(pdf_path, slide_number)
            if png_path:
                return png_path
        
        return None
    
    def _convert_pdf_to_png_pdf2image(self, pdf_path: str, slide_number: int) -> Optional[str]:
        """Convert PDF to PNG using pdf2image library."""
        
        try:
            from pdf2image import convert_from_path
            
            # Convert PDF pages to images (300 DPI for high quality)
            images = convert_from_path(pdf_path, dpi=300, fmt='PNG')
            
            if images:
                # Get the first page (single slide presentation)
                image = images[0]
                
                # Create temporary PNG file
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                    png_path = temp_file.name
                
                # Save the image
                image.save(png_path, 'PNG')
                return png_path
            
        except ImportError:
            # pdf2image not available
            pass
        except Exception:
            # Other errors
            pass
        
        return None
    
    def _convert_pdf_to_png_poppler(self, pdf_path: str, slide_number: int) -> Optional[str]:
        """Convert PDF to PNG using Poppler utilities (pdftoppm)."""
        
        try:
            # Create temporary directory for output
            with tempfile.TemporaryDirectory() as temp_dir:
                output_prefix = Path(temp_dir) / "slide"
                
                # Use pdftoppm to convert PDF to PNG
                cmd = [
                    "pdftoppm",
                    "-png",
                    "-r", "300",  # 300 DPI resolution
                    "-singlefile",  # Single file output
                    pdf_path,
                    str(output_prefix)
                ]
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                
                if result.returncode == 0:
                    # Find generated PNG file
                    png_files = list(Path(temp_dir).glob("*.png"))
                    if png_files:
                        source_file = png_files[0]
                        
                        # Create permanent temporary file
                        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                            png_path = temp_file.name
                        
                        # Copy the generated PNG
                        import shutil
                        shutil.copy2(source_file, png_path)
                        return png_path
                        
        except Exception:
            pass
        
        return None
    
    def _create_simple_fallback_thumbnail(self, pptx_path: str, slide_number: int) -> str:
        """
        Create a simple fallback thumbnail without external dependencies.
        
        This creates a basic colored square as a placeholder thumbnail.
        While not visually representative, it ensures the process continues.
        """
        # Create a simple 1x1 pixel PNG programmatically
        # This is a minimal PNG file that represents a colored square
        
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
            # Write a minimal PNG file (1x1 blue pixel)
            # PNG signature + IHDR + IDAT + IEND chunks
            png_data = bytes([
                # PNG signature
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                # IHDR chunk (13 bytes)
                0x00, 0x00, 0x00, 0x0D,  # Length
                0x49, 0x48, 0x44, 0x52,  # Type: IHDR
                0x00, 0x00, 0x00, 0x78,  # Width: 120
                0x00, 0x00, 0x00, 0x78,  # Height: 120  
                0x08, 0x02,              # Bit depth: 8, Color type: 2 (RGB)
                0x00, 0x00, 0x00,        # Compression, Filter, Interlace
                0x8D, 0xB8, 0xCB, 0x8F,  # CRC
                # IDAT chunk (minimal data for solid color)
                0x00, 0x00, 0x00, 0x16,  # Length
                0x49, 0x44, 0x41, 0x54,  # Type: IDAT
                0x78, 0x9C, 0xED, 0xC1, 0x01, 0x01, 0x00, 0x00, 
                0x00, 0x80, 0x90, 0xFE, 0x37, 0x96, 0x4E, 0x84,
                0x00, 0x02, 0x00, 0x00, 0x00, 0x01,
                0x24, 0x27, 0x0E, 0x1C,  # CRC
                # IEND chunk
                0x00, 0x00, 0x00, 0x00,  # Length
                0x49, 0x45, 0x4E, 0x44,  # Type: IEND
                0xAE, 0x42, 0x60, 0x82   # CRC
            ])
            
            temp_file.write(png_data)
            return temp_file.name
    
    def resize_thumbnail(self, thumbnail_path: str, target_height: int = None) -> str:
        """
        Resize a thumbnail to the target height while maintaining aspect ratio.
        Uses macOS built-in sips command for resizing.
        
        Args:
            thumbnail_path: Path to the source thumbnail
            target_height: Target height in pixels (defaults to config setting)
            
        Returns:
            Path to the resized thumbnail
        """
        if target_height is None:
            target_height = self.config['thumbnail_height']
        
        # Create output path for resized thumbnail
        input_path = Path(thumbnail_path)
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
            output_path = temp_file.name
        
        try:
            # Use macOS sips command to resize
            cmd = [
                'sips',
                '-Z', str(target_height),  # Resize maintaining aspect ratio
                thumbnail_path,
                '--out', output_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, timeout=10)
            
            if result.returncode == 0:
                return output_path
            else:
                # If sips fails, just copy the original
                import shutil
                shutil.copy2(thumbnail_path, output_path)
                return output_path
                
        except Exception:
            # If anything fails, just copy the original file
            import shutil
            shutil.copy2(thumbnail_path, output_path)
            return output_path
    
    def cleanup_temp_thumbnail(self, thumbnail_path: str) -> None:
        """Clean up a temporary thumbnail file."""
        try:
            if thumbnail_path and Path(thumbnail_path).exists():
                Path(thumbnail_path).unlink()
        except Exception:
            pass  # Ignore cleanup errors