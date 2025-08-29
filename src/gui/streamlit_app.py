"""
Streamlit web interface for Export for My Efficient Elements.

This module provides a user-friendly web interface for processing PowerPoint
presentations with real-time progress tracking and professional styling.
"""

import streamlit as st
import tempfile
import os
import base64
from pathlib import Path
import traceback

# Import configuration and core modules
from config.settings import get_gui_config, get_app_config, get_asset_path
from src.core.splitter import PowerPointSplitter


def get_base64_of_image(path):
    """Convert image to base64 string for HTML embedding."""
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()


def process_slides_with_progress(splitter, total_slides, progress_bar, status_text, progress_detail):
    """Process slides with detailed progress updates."""
    
    gui_config = get_gui_config()
    progress_colors = gui_config['progress_colors']
    
    # Create a progress callback function
    def progress_callback(current_slide, total_slides, slide_title, status):
        
        if current_slide <= total_slides:
            # Calculate progress (30% to 80% for slide processing)
            base_progress = 30 + (current_slide / total_slides) * 50
        else:
            # Final steps (80% to 100%)
            base_progress = 80 + ((current_slide - total_slides) / 2) * 20
            
        progress_bar.progress(int(base_progress))
        
        if status == "creating_pptx":
            # Update status
            status_text.text(f"📄 Creating PPTX {current_slide}/{total_slides}: {slide_title}")
            
            # Show PPTX creation progress
            progress_detail.markdown(f"""
            <div style="background-color: {progress_colors['creating_pptx']}; color: #FFFFFF; padding: 10px; border-radius: 5px; margin: 5px 0;">
            <strong>Slide {current_slide}/{total_slides}:</strong> {slide_title}<br>
            <small>📄 Creating individual PPTX file...</small>
            </div>
            """, unsafe_allow_html=True)
            
        elif status == "creating_thumbnail":
            # Update status
            status_text.text(f"🎨 Generating thumbnail {current_slide}/{total_slides}: {slide_title}")
            
            # Show thumbnail creation progress
            progress_detail.markdown(f"""
            <div style="background-color: {progress_colors['creating_thumbnail']}; color: #FFFFFF; padding: 10px; border-radius: 5px; margin: 5px 0;">
            <strong>Slide {current_slide}/{total_slides}:</strong> {slide_title}<br>
            <small>🎨 Generating high-quality thumbnail (this may take a moment)...</small>
            </div>
            """, unsafe_allow_html=True)
            
        elif status == "completed":
            # Show completion
            progress_detail.markdown(f"""
            <div style="background-color: {progress_colors['completed']}; color: #FFFFFF; padding: 10px; border-radius: 5px; margin: 5px 0;">
            <strong>Slide {current_slide}/{total_slides}:</strong> {slide_title}<br>
            <small>✅ PPTX and thumbnail created successfully</small>
            </div>
            """, unsafe_allow_html=True)
            
        elif status == "creating_xml":
            # Update status
            status_text.text("📄 Creating XML metadata...")
            
            # Show XML creation progress
            progress_detail.markdown(f"""
            <div style="background-color: {progress_colors['creating_xml']}; color: #FFFFFF; padding: 10px; border-radius: 5px; margin: 5px 0;">
            <strong>📄 XML Metadata:</strong><br>
            <small>Creating MyElements.xml with slide information...</small>
            </div>
            """, unsafe_allow_html=True)
            
        elif status == "creating_zip":
            # Update status
            status_text.text("📦 Creating zip archive...")
            
            # Show zip creation progress
            progress_detail.markdown(f"""
            <div style="background-color: {progress_colors['creating_zip']}; color: #FFFFFF; padding: 10px; border-radius: 5px; margin: 5px 0;">
            <strong>📦 Final Archive:</strong><br>
            <small>Compressing all files into downloadable zip archive...</small>
            </div>
            """, unsafe_allow_html=True)
            
        elif status == "export_complete":
            # Update status
            status_text.text("✅ Export completed successfully!")
            
            # Show final completion
            progress_detail.markdown(f"""
            <div style="background-color: {progress_colors['export_complete']}; color: #FFFFFF; padding: 10px; border-radius: 5px; margin: 5px 0;">
            <strong>🎉 Export Complete!</strong><br>
            <small>All elements exported successfully and ready for download</small>
            </div>
            """, unsafe_allow_html=True)
    
    # Do the actual processing with real-time progress
    status_text.text("⚡ Starting slide processing...")
    created_files = splitter.split_slides(progress_callback=progress_callback)
    
    return created_files


def main():
    """Main Streamlit application."""
    
    # Load configurations
    gui_config = get_gui_config()
    app_config = get_app_config()
    
    # Configure page
    st.set_page_config(
        page_title=gui_config['title'],
        page_icon=gui_config['icon'],
        layout=gui_config['layout'],
        initial_sidebar_state=gui_config['sidebar_state']
    )
    
    # Custom CSS for better styling
    colors = gui_config['colors']
    st.markdown(f"""
    <style>
    .main-header {{
        text-align: center;
        color: {colors['primary']};
        margin-bottom: 2rem;
    }}
    .feature-box {{
        background-color: {colors['dark']};
        color: #FFFFFF;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid {colors['info']};
        margin: 1rem 0;
    }}
    .success-box {{
        background-color: {colors['success']};
        color: #FFFFFF;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid {colors['success']};
        margin: 1rem 0;
    }}
    .stApp > div:first-child > div:first-child > div:first-child {{
        padding-top: 2rem;
    }}
    .block-container {{
        max-width: 1200px;
        padding-left: 2rem;
        padding-right: 2rem;
    }}
    </style>
    """, unsafe_allow_html=True)
    
    # Header with logo - perfectly centered
    logo_path = get_asset_path("EfficientElementsLogo.png")
    st.markdown(
        f"""
        <div style="display: flex; justify-content: center; align-items: center; margin: 1rem 0;">
            <img src="data:image/png;base64,{get_base64_of_image(logo_path)}" width="{gui_config['logo_width']}">
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Title
    st.markdown(f'<h1 class="main-header">{app_config["name"]}</h1>', unsafe_allow_html=True)
    st.markdown(f'<p style="text-align: center; color: {colors["secondary"]};">{app_config["description"]}</p>', unsafe_allow_html=True)
    
    # Main content area - wider layout
    col1, col2, col3 = st.columns([1, 6, 1])
    
    with col2:
        # File upload section
        st.markdown("### 📁 Upload PowerPoint File")
        uploaded_file = st.file_uploader(
            "Choose a PowerPoint file",
            type=app_config['supported_types'],
            help=f"Upload your PowerPoint presentation ({', '.join([f'.{ext}' for ext in app_config['supported_types']])} format)"
        )
        
        # Folder name input
        st.markdown("### 📁 Folder name")
        group_name = st.text_input(
            "Name",
            value="My Presentation",
            help="This name will be used for organizing your exported elements"
        )
        
        # Export button
        if st.button("🚀 Export now", type="primary", use_container_width=True):
            if uploaded_file is not None and group_name.strip():
                process_powerpoint(uploaded_file, group_name.strip())
            elif uploaded_file is None:
                st.error("Please upload a PowerPoint file first!")
            else:
                st.error("Please provide a folder name!")
        
        # Show file info if uploaded
        if uploaded_file is not None:
            st.markdown("### 📋 File Information")
            file_size_mb = uploaded_file.size / 1024 / 1024
            st.markdown(f"""
            <div class="feature-box">
            <strong>📄 File:</strong> {uploaded_file.name}<br>
            <strong>📏 Size:</strong> {file_size_mb:.1f} MB
            </div>
            """, unsafe_allow_html=True)


def process_powerpoint(uploaded_file, group_name):
    """Process the uploaded PowerPoint file."""
    
    # Create progress containers
    progress_container = st.container()
    result_container = st.container()
    
    with progress_container:
        st.markdown("### 🔄 Processing...")
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # Save uploaded file to temporary location
            status_text.text("📁 Saving uploaded file...")
            progress_bar.progress(10)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{uploaded_file.name}") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                temp_input_path = tmp_file.name
            
            # Initialize the splitter
            status_text.text("🚀 Initializing PowerPoint Splitter...")
            progress_bar.progress(20)
            
            # Extract original filename without extension for clean zip naming
            original_base_name = Path(uploaded_file.name).stem
            
            splitter = PowerPointSplitter(
                input_file=temp_input_path,
                group_name=group_name,
                base_name=original_base_name
            )
            
            # Get presentation info first
            from pptx import Presentation
            prs = Presentation(temp_input_path)
            total_slides = len(prs.slides)
            
            status_text.text(f"📊 Found {total_slides} slides to process...")
            progress_bar.progress(30)
            
            # Process slides with detailed progress
            progress_detail = st.empty()
            
            # Create a custom processing with progress updates
            created_files = process_slides_with_progress(
                splitter, 
                total_slides, 
                progress_bar, 
                status_text, 
                progress_detail
            )
            
            progress_bar.progress(90)
            status_text.text("📦 Finalizing zip archive...")
            
            # Find the created zip file using the original base name
            input_path = Path(temp_input_path)
            
            # Look for zip files in the same directory as the temp input using the original filename
            zip_files = list(input_path.parent.glob(f"{original_base_name}_*.zip"))
            
            if zip_files:
                # Get the most recent zip file
                zip_file_path = max(zip_files, key=lambda x: x.stat().st_mtime)
                
                progress_bar.progress(100)
                status_text.text("✅ Processing complete!")
                
                # Clean up temp input file
                try:
                    os.unlink(temp_input_path)
                except:
                    pass
                
                # Show success and provide download
                show_success_result(zip_file_path, group_name)
                
            else:
                st.error("❌ No zip file was created. Please check the processing details above.")
                
        except Exception as e:
            progress_bar.progress(0)
            status_text.text("❌ Processing failed!")
            st.error(f"Error during processing: {str(e)}")
            if st.checkbox("Show detailed error information"):
                st.code(traceback.format_exc())
            
            # Clean up temp file on error
            try:
                if 'temp_input_path' in locals():
                    os.unlink(temp_input_path)
            except:
                pass


def show_success_result(zip_file_path, group_name):
    """Display success message and provide download link."""
    
    gui_config = get_gui_config()
    colors = gui_config['colors']
    
    st.markdown("---")
    
    # Success message
    st.markdown(f"""
    <div class="success-box">
    <h3>🎉 Export Complete!</h3>
    <p>Your presentation has been successfully converted into importable My Efficient Elements.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # File information
    zip_size = zip_file_path.stat().st_size / 1024 / 1024  # Size in MB
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Export Results")
        st.markdown(f"""
        - **📦 Archive:** {zip_file_path.name}
        - **📏 Size:** {zip_size:.1f} MB
        - **📁 Folder:** {group_name}
        """)
    
    with col2:
        st.markdown("### 📥 Download")
        
        # Read the zip file for download
        with open(zip_file_path, 'rb') as f:
            zip_data = f.read()
        
        st.download_button(
            label="📁 Download My Elements Export",
            data=zip_data,
            file_name=zip_file_path.name,
            mime="application/zip",
            type="primary",
            use_container_width=True
        )
    

    
    # Success tips
    st.markdown("### 💡 What's Next?")
    st.markdown("""
    - Open PowerPoint
    - Open element wizard by clicking on Bugs or Icons button
    - Go to My elements in the bottom of left panel and select import button from the bottom
    - Use the downloaded zip file to import these elements
    """)


if __name__ == "__main__":
    main()
