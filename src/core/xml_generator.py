"""
XML metadata generation module for Export for My Efficient Elements.

This module handles the creation of MyElements.xml files with proper
structure and formatting for importing into presentation software.
"""

import xml.etree.ElementTree as ET
import xml.dom.minidom
from pathlib import Path
from typing import List, Dict, Any

from config.settings import get_processing_config
from src.utils.uuid_utils import generate_reproducible_uuid


class XMLGenerator:
    """Handles XML metadata file generation."""
    
    def __init__(self):
        self.config = get_processing_config()
    
    def create_xml_metadata(
        self, 
        group_name: str, 
        slide_metadata: List[Dict[str, str]], 
        output_dir: Path
    ) -> Path:
        """
        Create XML metadata file for the exported slides.
        
        Args:
            group_name: Name of the group for the XML metadata
            slide_metadata: List of slide metadata dictionaries
            output_dir: Directory to save the XML file
            
        Returns:
            Path to the created XML file
        """
        # Create the root element
        root = ET.Element("ee4p")
        
        # Generate reproducible UUID for the group
        group_id = generate_reproducible_uuid(group_name)
        
        # Create group element
        group_element = ET.SubElement(root, "group")
        group_element.set("id", group_id)
        group_element.set("name", group_name)
        
        # Add individual slide elements
        for slide_data in slide_metadata:
            element = ET.SubElement(group_element, "element")
            element.set("name", slide_data['name'])
            element.set("thumbMode", slide_data['thumbMode'])
            element.set("id", slide_data['id'])
        
        # Generate the XML file
        xml_filename = self.config['xml_filename']
        xml_path = output_dir / xml_filename
        
        # Create formatted XML
        self._write_formatted_xml(root, xml_path)
        
        return xml_path
    
    def _write_formatted_xml(self, root: ET.Element, output_path: Path) -> None:
        """
        Write XML with proper formatting and without XML declaration.
        
        Args:
            root: Root XML element
            output_path: Path to write the XML file
        """
        # Convert to string with minidom for pretty formatting
        rough_string = ET.tostring(root, encoding='unicode')
        parsed = xml.dom.minidom.parseString(rough_string)
        
        # Get pretty-printed XML
        pretty_xml = parsed.documentElement.toprettyxml(indent="  ")
        
        # Remove the XML declaration line and empty lines
        lines = pretty_xml.split('\n')
        # Skip empty lines and keep only content lines
        content_lines = [line for line in lines if line.strip()]
        
        # Write to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(content_lines))
    
    def validate_xml_structure(self, xml_path: Path) -> tuple[bool, str]:
        """
        Validate the structure of a generated XML file.
        
        Args:
            xml_path: Path to the XML file to validate
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Check root element
            if root.tag != "ee4p":
                return False, "Root element must be 'ee4p'"
            
            # Check for group element
            group = root.find("group")
            if group is None:
                return False, "Missing 'group' element"
            
            # Check group attributes
            required_group_attrs = ["id", "name"]
            for attr in required_group_attrs:
                if attr not in group.attrib:
                    return False, f"Missing required group attribute: {attr}"
            
            # Check element children
            elements = group.findall("element")
            if not elements:
                return False, "No 'element' children found in group"
            
            # Check element attributes
            required_element_attrs = ["name", "thumbMode", "id"]
            for i, element in enumerate(elements):
                for attr in required_element_attrs:
                    if attr not in element.attrib:
                        return False, f"Missing required attribute '{attr}' in element {i+1}"
            
            return True, "XML structure is valid"
            
        except ET.ParseError as e:
            return False, f"XML parsing error: {e}"
        except Exception as e:
            return False, f"Validation error: {e}"
    
    def create_sample_xml(self, output_path: Path) -> Path:
        """
        Create a sample XML file for reference.
        
        Args:
            output_path: Path where to create the sample file
            
        Returns:
            Path to the created sample file
        """
        # Sample data
        sample_group_name = "Sample Presentation"
        sample_slides = [
            {"name": "Introduction", "thumbMode": "1", "id": "12345678-1234-5678-1234-123456789abc"},
            {"name": "Overview", "thumbMode": "1", "id": "87654321-4321-8765-4321-cba987654321"},
            {"name": "Conclusion", "thumbMode": "1", "id": "11111111-2222-3333-4444-555555555555"}
        ]
        
        return self.create_xml_metadata(sample_group_name, sample_slides, output_path.parent)
    
    def extract_metadata_from_xml(self, xml_path: Path) -> Dict[str, Any]:
        """
        Extract metadata from an existing XML file.
        
        Args:
            xml_path: Path to the XML file
            
        Returns:
            Dictionary containing extracted metadata
        """
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            group = root.find("group")
            if group is None:
                return {}
            
            metadata = {
                "group_id": group.get("id"),
                "group_name": group.get("name"),
                "elements": []
            }
            
            # Extract element information
            for element in group.findall("element"):
                element_data = {
                    "name": element.get("name"),
                    "thumbMode": element.get("thumbMode"),
                    "id": element.get("id")
                }
                metadata["elements"].append(element_data)
            
            return metadata
            
        except Exception as e:
            print(f"Error extracting metadata: {e}")
            return {}
    
    def update_xml_metadata(
        self, 
        xml_path: Path, 
        new_group_name: str = None, 
        additional_elements: List[Dict[str, str]] = None
    ) -> bool:
        """
        Update an existing XML metadata file.
        
        Args:
            xml_path: Path to the XML file to update
            new_group_name: New group name (optional)
            additional_elements: Additional elements to add (optional)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            group = root.find("group")
            if group is None:
                return False
            
            # Update group name if provided
            if new_group_name:
                group.set("name", new_group_name)
                # Regenerate group ID for consistency
                new_group_id = generate_reproducible_uuid(new_group_name)
                group.set("id", new_group_id)
            
            # Add additional elements if provided
            if additional_elements:
                for element_data in additional_elements:
                    element = ET.SubElement(group, "element")
                    element.set("name", element_data['name'])
                    element.set("thumbMode", element_data.get('thumbMode', '1'))
                    element.set("id", element_data['id'])
            
            # Write updated XML
            self._write_formatted_xml(root, xml_path)
            return True
            
        except Exception as e:
            print(f"Error updating XML metadata: {e}")
            return False
