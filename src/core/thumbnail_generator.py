"""
Thumbnail generation module for Export for My Efficient Elements.

This module handles thumbnail generation using macOS Quick Look (qlmanage)
with a simple fallback for cases where Quick Look is not available.
Optimized for macOS systems with no external dependencies.
"""

import os
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional

from config.settings import get_processing_config


class SlideThumbnailGenerator:
    """Streamlined thumbnail generator optimized for macOS Quick Look."""
    
    def __init__(self):
        self.config = get_processing_config()
        
        # Check available conversion methods
        self.conversion_methods = self._detect_conversion_methods()
        if self.config['verbose']:
            print(f"🔍 Available conversion methods: {', '.join(self.conversion_methods)}")
    
    def _detect_conversion_methods(self) -> List[str]:
        """Detect which conversion methods are available on this system."""
        methods = []
        
        # Check for macOS Quick Look (qlmanage)
        try:
            result = subprocess.run(['qlmanage', '-h'], capture_output=True, timeout=3)
            if result.returncode == 0:
                methods.append('qlmanage')
        except:
            pass
        
        # Always available simple fallback
        methods.append('simple_fallback')
        
        return methods
    
    def create_high_quality_thumbnail_from_pptx(self, pptx_path: str, slide_number: int) -> Optional[str]:
        """
        Create a high-quality thumbnail from PPTX file using macOS Quick Look.
        
        Args:
            pptx_path: Path to the PPTX file
            slide_number: Slide number for identification
            
        Returns:
            Path to the generated thumbnail file, or None if failed
        """
        if self.config['verbose']:
            print(f"    🎨 Generating high-quality thumbnail using macOS Quick Look...")
        
        # Try qlmanage first (primary method)
        if 'qlmanage' in self.conversion_methods:
            thumbnail_path = self._convert_with_qlmanage(pptx_path)
            if thumbnail_path:
                if self.config['verbose']:
                    print(f"    ✅ Used macOS Quick Look for high-quality conversion")
                return thumbnail_path
        
        # Fallback to simple placeholder
        if self.config['verbose']:
            print(f"    ℹ️  Using simple fallback thumbnail")
        return self._create_simple_fallback_thumbnail(pptx_path, slide_number)
    
    def _convert_with_qlmanage(self, pptx_path: str) -> Optional[str]:
        """Convert PPTX to image using macOS Quick Look."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Use qlmanage to generate thumbnail
            cmd = [
                'qlmanage', 
                '-t', str(pptx_path),
                '-s', '1440',  # Size for high quality
                '-o', temp_dir
            ]
            
            try:
                timeout = self.config['thumbnail_methods']['macos_quicklook']['timeout']
                result = subprocess.run(cmd, capture_output=True, timeout=timeout)
                
                if result.returncode == 0:
                    # qlmanage creates files with specific naming
                    generated_files = list(Path(temp_dir).glob("*.png"))
                    if generated_files:
                        # Move the generated file to a permanent location
                        source_file = generated_files[0]
                        # Create a temporary file for the thumbnail
                        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                            temp_thumbnail_path = temp_file.name
                        
                        # Copy the generated thumbnail
                        import shutil
                        shutil.copy2(source_file, temp_thumbnail_path)
                        return temp_thumbnail_path
                        
            except subprocess.TimeoutExpired:
                if self.config['verbose']:
                    print(f"    ⚠️  qlmanage conversion timed out")
            except Exception as e:
                if self.config['verbose']:
                    print(f"    ⚠️  qlmanage conversion failed: {e}")
            
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