"""
File utility functions for Export for My Efficient Elements.

This module provides utilities for file operations, archive creation,
and cleanup tasks used throughout the application.
"""

import os
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import List, Tuple, Optional

from config.settings import get_processing_config


def create_temp_directory(prefix: Optional[str] = None) -> Path:
    """
    Create a temporary directory for processing.
    
    Args:
        prefix: Optional prefix for the directory name
        
    Returns:
        Path object pointing to the created directory
    """
    import tempfile
    
    config = get_processing_config()
    dir_prefix = prefix or config['temp_prefix']
    
    temp_dir = Path(tempfile.mkdtemp(prefix=dir_prefix))
    return temp_dir


def ensure_directory_exists(directory_path: Path) -> None:
    """
    Ensure a directory exists, creating it if necessary.
    
    Args:
        directory_path: Path to the directory
    """
    directory_path.mkdir(parents=True, exist_ok=True)


def get_file_size_mb(file_path: Path) -> float:
    """
    Get file size in megabytes.
    
    Args:
        file_path: Path to the file
        
    Returns:
        File size in MB, rounded to 1 decimal place
    """
    if not file_path.exists():
        return 0.0
    
    size_bytes = file_path.stat().st_size
    size_mb = size_bytes / (1024 * 1024)
    return round(size_mb, 1)


def is_supported_file_type(file_path: Path, supported_types: List[str]) -> bool:
    """
    Check if file type is supported.
    
    Args:
        file_path: Path to the file
        supported_types: List of supported file extensions
        
    Returns:
        True if file type is supported, False otherwise
    """
    if not file_path.exists():
        return False
    
    file_extension = file_path.suffix.lower().lstrip('.')
    return file_extension in [ext.lower() for ext in supported_types]


def generate_timestamped_filename(base_name: str, extension: str = "zip") -> str:
    """
    Generate a filename with timestamp.
    
    Args:
        base_name: Base name for the file
        extension: File extension (without dot)
        
    Returns:
        Timestamped filename
    """
    config = get_processing_config()
    timestamp = datetime.now().strftime(config['timestamp_format'])
    return f"{base_name}_{timestamp}.{extension}"


def create_zip_archive(
    files_to_compress: List[Path], 
    output_path: Path, 
    exclude_system_files: bool = True
) -> Tuple[bool, int, float]:
    """
    Create a zip archive from a list of files.
    
    Args:
        files_to_compress: List of file paths to include in the zip
        output_path: Path where the zip file should be created
        exclude_system_files: Whether to exclude system files (e.g., .DS_Store)
        
    Returns:
        Tuple of (success, file_count, archive_size_mb)
    """
    try:
        # System files to exclude
        system_files = {'.DS_Store', 'Thumbs.db', '.git', '.gitignore'} if exclude_system_files else set()
        
        files_added = 0
        
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in files_to_compress:
                if file_path.exists() and file_path.name not in system_files:
                    # Add file at root level (no directory structure)
                    zipf.write(file_path, file_path.name)
                    files_added += 1
        
        # Get archive size
        archive_size_mb = get_file_size_mb(output_path)
        
        return True, files_added, archive_size_mb
        
    except Exception as e:
        print(f"Error creating zip archive: {e}")
        return False, 0, 0.0


def cleanup_files(file_paths: List[Path], verbose: bool = True) -> int:
    """
    Clean up a list of files.
    
    Args:
        file_paths: List of file paths to delete
        verbose: Whether to print verbose output
        
    Returns:
        Number of files successfully deleted
    """
    deleted_count = 0
    
    for file_path in file_paths:
        try:
            if file_path.exists():
                if file_path.is_file():
                    file_path.unlink()
                elif file_path.is_dir():
                    shutil.rmtree(file_path, ignore_errors=True)
                deleted_count += 1
                
        except Exception as e:
            if verbose:
                print(f"Warning: Could not delete {file_path}: {e}")
    
    return deleted_count


def cleanup_directory(directory_path: Path, remove_directory: bool = True, verbose: bool = True) -> bool:
    """
    Clean up a directory and optionally remove it.
    
    Args:
        directory_path: Path to the directory
        remove_directory: Whether to remove the directory itself
        verbose: Whether to print verbose output
        
    Returns:
        True if successful, False otherwise
    """
    try:
        if not directory_path.exists():
            return True
        
        if directory_path.is_dir():
            if remove_directory:
                shutil.rmtree(directory_path, ignore_errors=True)
                if verbose:
                    print(f"    🗑️  Removed directory: {directory_path.name}")
            else:
                # Just clean contents
                for item in directory_path.iterdir():
                    if item.is_file():
                        item.unlink()
                    elif item.is_dir():
                        shutil.rmtree(item, ignore_errors=True)
                        
                # Remove directory if empty
                try:
                    directory_path.rmdir()
                    if verbose:
                        print(f"    🗑️  Removed empty directory: {directory_path.name}")
                except OSError:
                    # Directory not empty, leave it
                    pass
            
            return True
            
    except Exception as e:
        if verbose:
            print(f"Warning: Could not clean directory {directory_path}: {e}")
        return False
    
    return False


def get_files_in_directory(directory_path: Path, pattern: str = "*") -> List[Path]:
    """
    Get all files in a directory matching a pattern.
    
    Args:
        directory_path: Path to the directory
        pattern: Glob pattern to match files
        
    Returns:
        List of matching file paths
    """
    if not directory_path.exists() or not directory_path.is_dir():
        return []
    
    return list(directory_path.glob(pattern))


def copy_file_with_new_name(source_path: Path, destination_dir: Path, new_name: str) -> Path:
    """
    Copy a file to a new location with a new name.
    
    Args:
        source_path: Source file path
        destination_dir: Destination directory
        new_name: New filename (including extension)
        
    Returns:
        Path to the copied file
    """
    ensure_directory_exists(destination_dir)
    destination_path = destination_dir / new_name
    shutil.copy2(source_path, destination_path)
    return destination_path


def validate_file_access(file_path: Path) -> Tuple[bool, str]:
    """
    Validate that a file exists and is accessible.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    if not file_path.exists():
        return False, f"File does not exist: {file_path}"
    
    if not file_path.is_file():
        return False, f"Path is not a file: {file_path}"
    
    if not os.access(file_path, os.R_OK):
        return False, f"File is not readable: {file_path}"
    
    return True, ""
