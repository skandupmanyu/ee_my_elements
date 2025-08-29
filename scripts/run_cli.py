#!/usr/bin/env python3
"""
Command-line interface for Export for My Efficient Elements.

This script provides the terminal interface for processing PowerPoint
presentations with full automation and debugging capabilities.
"""

import sys
import argparse
from pathlib import Path

# Add the project root to the Python path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from config.settings import get_app_config, SUPPORTED_FILE_TYPES
from src.core.splitter import PowerPointSplitter


def main():
    """Main function to handle command-line interface."""
    
    app_config = get_app_config()
    
    parser = argparse.ArgumentParser(
        description=f"{app_config['name']} - {app_config['description']}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
  python scripts/run_cli.py presentation.pptx
  python scripts/run_cli.py presentation.pptx -o my_slides/
  python scripts/run_cli.py presentation.pptx -g "PodHandler"
  python scripts/run_cli.py presentation.pptx --group-name "My Custom Group" --output-dir ./individual_slides/
  
Supported file types: {', '.join([f'.{ext}' for ext in SUPPORTED_FILE_TYPES])}
Max file size: {app_config['max_file_size_mb']} MB
        """
    )
    
    parser.add_argument(
        "input_file",
        help=f"Path to the input PowerPoint file ({', '.join([f'.{ext}' for ext in SUPPORTED_FILE_TYPES])})"
    )
    
    parser.add_argument(
        "-o", "--output-dir",
        help="Directory to save the individual slide files (default: temporary directory)",
        default=None
    )
    
    parser.add_argument(
        "-g", "--group-name",
        help="Name of the group for XML metadata (default: derived from presentation filename)",
        default=None
    )
    
    parser.add_argument(
        "-b", "--base-name",
        help="Base name for the output zip file (default: input filename)",
        default=None
    )
    
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose output"
    )
    
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug mode with detailed error information"
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version=f"{app_config['name']} v1.0.0"
    )
    
    args = parser.parse_args()
    
    try:
        # Validate input file
        input_path = Path(args.input_file)
        if not input_path.exists():
            print(f"❌ Error: Input file not found: {input_path}")
            sys.exit(1)
        
        if not input_path.suffix.lower() in [f'.{ext}' for ext in SUPPORTED_FILE_TYPES]:
            supported_ext = ', '.join([f'.{ext}' for ext in SUPPORTED_FILE_TYPES])
            print(f"❌ Error: Input file must be a PowerPoint file ({supported_ext})")
            sys.exit(1)
        
        # Check file size
        file_size_mb = input_path.stat().st_size / (1024 * 1024)
        if file_size_mb > app_config['max_file_size_mb']:
            print(f"❌ Error: File size ({file_size_mb:.1f} MB) exceeds maximum allowed size ({app_config['max_file_size_mb']} MB)")
            sys.exit(1)
        
        # Determine group name
        group_name = args.group_name
        if not group_name:
            # Use the presentation filename (without extension) as default group name
            group_name = input_path.stem
        
        # Create the splitter
        splitter = PowerPointSplitter(
            input_file=str(input_path),
            output_dir=args.output_dir,
            group_name=group_name,
            base_name=args.base_name
        )
        
        # Split the slides
        created_files = splitter.split_slides()
        
        print(f"\n✅ Successfully created {len(created_files)} individual slide files")
        
        if args.verbose:
            print(f"📁 Working directory: {splitter.output_dir}")
            print(f"🏷️  Group name: {group_name}")
            print(f"📦 Base name: {splitter.base_name}")
            print("\nCreated files:")
            for file_path in created_files:
                print(f"  • {Path(file_path).name}")
        
    except FileNotFoundError as e:
        print(f"❌ Error: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"❌ Error: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print(f"\n⚠️  Process interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        if args.debug:
            import traceback
            print("\nDetailed error information:")
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
