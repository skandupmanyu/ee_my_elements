"""
UUID utility functions for Export for My Efficient Elements.

This module provides utilities for generating reproducible and unique UUIDs
for consistent identification across multiple runs.
"""

import uuid
from typing import Optional


def generate_reproducible_uuid(input_string: str, namespace: Optional[str] = None) -> str:
    """
    Generate a reproducible UUID based on an input string.
    
    This ensures that the same input string will always generate the same UUID,
    which is important for consistent group identification across multiple exports.
    
    Args:
        input_string: The string to generate UUID from (e.g., group name)
        namespace: Optional namespace string for additional uniqueness
        
    Returns:
        String representation of the generated UUID
    """
    # Use a fixed namespace for reproducibility
    if namespace is None:
        # Fixed namespace UUID for the project
        namespace_uuid = uuid.UUID('12345678-1234-5678-1234-123456789abc')
    else:
        # Generate namespace UUID from provided string
        namespace_uuid = uuid.uuid5(uuid.NAMESPACE_DNS, namespace)
    
    # Generate UUID5 based on the input string and namespace
    generated_uuid = uuid.uuid5(namespace_uuid, input_string)
    
    return str(generated_uuid)


def generate_unique_uuid() -> str:
    """
    Generate a completely unique UUID for individual slide identification.
    
    This creates a new UUID4 for each call, ensuring uniqueness for
    individual slide files that should not be reproducible.
    
    Returns:
        String representation of the generated unique UUID
    """
    return str(uuid.uuid4())


def is_valid_uuid(uuid_string: str) -> bool:
    """
    Validate if a string is a valid UUID format.
    
    Args:
        uuid_string: String to validate
        
    Returns:
        True if valid UUID format, False otherwise
    """
    try:
        uuid.UUID(uuid_string)
        return True
    except ValueError:
        return False


def format_uuid_for_filename(uuid_string: str) -> str:
    """
    Format UUID string to be safe for use in filenames.
    
    Args:
        uuid_string: UUID string to format
        
    Returns:
        Filename-safe UUID string
    """
    if not is_valid_uuid(uuid_string):
        raise ValueError(f"Invalid UUID format: {uuid_string}")
    
    # UUIDs are already filename-safe, but this function exists
    # for consistency and potential future enhancements
    return uuid_string.lower()


def create_group_metadata(group_name: str, project_namespace: str = "my-efficient-elements") -> dict:
    """
    Create group metadata with reproducible UUID and formatting.
    
    Args:
        group_name: Name of the group
        project_namespace: Project-specific namespace for UUID generation
        
    Returns:
        Dictionary containing group metadata
    """
    group_id = generate_reproducible_uuid(group_name, project_namespace)
    
    return {
        'name': group_name,
        'id': group_id,
        'type': 'group',
        'namespace': project_namespace
    }


def create_element_metadata(element_name: str, thumb_mode: str = "1") -> dict:
    """
    Create element metadata with unique UUID.
    
    Args:
        element_name: Name of the element (slide title)
        thumb_mode: Thumbnail mode setting
        
    Returns:
        Dictionary containing element metadata
    """
    element_id = generate_unique_uuid()
    
    return {
        'name': element_name,
        'id': element_id,
        'thumbMode': thumb_mode,
        'type': 'element'
    }
