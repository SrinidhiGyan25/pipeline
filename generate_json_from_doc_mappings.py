#!/usr/bin/env python3
"""
Generate JSON mapping file from Word document for PPTImageInserter
Extracts slide-to-image mappings from .docx files and converts them to JSON format
with collision detection and automatic position assignment
"""

import os
import json
import re
from pathlib import Path
from typing import Dict, List, Optional, Set
from enum import Enum
from docx import Document

class Position(Enum):
    """Position enumeration for consistent position management"""
    TOP_LEFT = "top-left"
    TOP_RIGHT = "top-right"
    BOTTOM_LEFT = "bottom-left"
    BOTTOM_RIGHT = "bottom-right"
    CENTER = "center"
    CUSTOM = "custom"

class PositionManager:
    """Manages position assignment with collision detection"""
    
    def __init__(self):
        self.available_positions = [
            Position.BOTTOM_LEFT,
            Position.BOTTOM_RIGHT,
            Position.TOP_RIGHT
        ]
        # Track occupied positions per slide
        self.occupied_positions: Dict[int, Set[Position]] = {}
    
    def get_next_available_position(self, slide_number: int) -> Optional[Position]:
        """Get next available position on a slide"""
        if slide_number not in self.occupied_positions:
            self.occupied_positions[slide_number] = set()
        
        for position in self.available_positions:
            if position not in self.occupied_positions[slide_number]:
                return position
        return None
    
    def occupy_position(self, slide_number: int, position: Position):
        """Mark position as occupied"""
        if slide_number not in self.occupied_positions:
            self.occupied_positions[slide_number] = set()
        self.occupied_positions[slide_number].add(position)
    
    def is_slide_full(self, slide_number: int) -> bool:
        """Check if all positions on slide are occupied"""
        return self.get_next_available_position(slide_number) is None
    
    def get_available_count(self, slide_number: int) -> int:
        """Get count of available positions on slide"""
        if slide_number not in self.occupied_positions:
            return len(self.available_positions)
        return len(self.available_positions) - len(self.occupied_positions[slide_number])

def extract_mappings_from_docx(docx_path: str) -> Dict[str, List[str]]:
    """
    Extract slide-to-image mappings from a Word document
    
    Expected format in document:
    - "Slide 1 -- image1.jpg, image2.png"
    - "Slide 2 -- image3.jpg"
    - "Slide 10 -- image4.png, image5.jpg, image6.gif"
    
    Args:
        docx_path: Path to the Word document
        
    Returns:
        Dict mapping slide numbers to image lists: {"slide_1": ["image1.jpg", "image2.png"], ...}
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Document file not found: {docx_path}")
    
    print(f"📄 Reading mappings from: {docx_path}")
    
    try:
        doc = Document(docx_path)
        full_text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
    except Exception as e:
        raise ValueError(f"Error reading Word document: {e}")
    
    mappings = {}
    
    # Pattern to match: "Slide <number> -- <images>"
    # More flexible pattern that handles various formats
    slide_patterns = [
        r'Slide\s+(\d+)\s*--\s*(.+)',           # "Slide 1 -- image1.jpg, image2.png"
        r'Slide\s+(\d+)\s*:\s*(.+)',            # "Slide 1: image1.jpg, image2.png"
        r'Slide\s+(\d+)\s*-\s*(.+)',            # "Slide 1 - image1.jpg, image2.png"
        r'(\d+)\s*--\s*(.+)',                   # "1 -- image1.jpg, image2.png"
        r'(\d+)\s*:\s*(.+)',                    # "1: image1.jpg, image2.png"
    ]
    
    for pattern in slide_patterns:
        matches = re.findall(pattern, full_text, re.MULTILINE | re.IGNORECASE)
        if matches:
            print(f"✓ Found {len(matches)} slide mappings using pattern: {pattern}")
            break
    
    if not matches:
        print("⚠️  No slide mappings found. Expected format examples:")
        print("   - Slide 1 -- image1.jpg, image2.png")
        print("   - Slide 2 -- image3.jpg")
        print("   - Slide 10 -- image4.png, image5.jpg")
        return {}
    
    # Process each match
    for slide_num, images_str in matches:
        slide_key = f"slide_{slide_num}"
        
        # Split images by comma and clean up
        raw_images = [img.strip() for img in images_str.split(',')]
        clean_images = []
        
        for img in raw_images:
            # Remove any extra whitespace and filter out empty strings
            img = img.strip()
            if img:
                # Remove any quotes or extra characters
                img = img.strip('"\'')
                clean_images.append(img)
        
        if clean_images:
            mappings[slide_key] = clean_images
            print(f"  Slide {slide_num}: {clean_images}")
    
    print(f"✓ Extracted {len(mappings)} slide mappings")
    return mappings

def convert_to_ppt_format_with_collision_detection(slide_mappings: Dict[str, List[str]], 
                                                 default_width: float = 6.0, 
                                                 default_height: float = 4.0,
                                                 default_position: str = "center") -> List[Dict]:
    """
    Convert slide-based mappings to the format expected by PPTImageInserter
    with collision detection and automatic position assignment
    
    Args:
        slide_mappings: {"slide_1": ["image1.jpg", "image2.png"], ...}
        default_width: Default image width in inches
        default_height: Default image height in inches
        default_position: Default position ("center", "top-left", etc.)
        
    Returns:
        List of mappings in PPTImageInserter format with collision detection
    """
    position_manager = PositionManager()
    ppt_mappings = []
    image_counter = 1
    
    # Sort slides by number to ensure consistent ordering
    sorted_slides = sorted(slide_mappings.items(), 
                          key=lambda x: int(x[0].split('_')[1]))
    
    print(f"\n=== Converting to PPT Format with Collision Detection ===")
    
    # First pass: analyze slide usage
    print("\n=== Slide Usage Analysis ===")
    max_positions = len(position_manager.available_positions)
    
    for slide_key, image_names in sorted_slides:
        slide_num = int(slide_key.split('_')[1])
        image_count = len(image_names)
        
        if image_count > max_positions:
            print(f"Slide {slide_num}: {image_count} images (exceeds {max_positions} positions)")
        else:
            print(f"Slide {slide_num}: {image_count} images")
    
    # Second pass: assign positions with collision detection
    print(f"\n=== Position Assignment ===")
    
    for slide_key, image_names in sorted_slides:
        slide_num = int(slide_key.split('_')[1])
        
        print(f"\n📍 Processing Slide {slide_num} with {len(image_names)} images")
        
        # Check if slide can accommodate all images
        available_positions = position_manager.get_available_count(slide_num)
        
        if len(image_names) > available_positions:
            print(f"  ⚠️ Warning: Slide {slide_num} needs {len(image_names)} positions but only has {available_positions} available")
            print(f"  ✓ Will distribute excess images to new slides")
        
        current_slide = slide_num
        
        for i, image_name in enumerate(image_names):
            # Check if current slide has space
            if position_manager.is_slide_full(current_slide):
                # Move to next slide
                current_slide = _find_next_available_slide(position_manager, current_slide)
                print(f"  ✓ Moving to slide {current_slide} for image {image_counter}")
            
            # Get next available position
            position = position_manager.get_next_available_position(current_slide)
            
            if position:
                position_manager.occupy_position(current_slide, position)
                
                # Create mapping entry
                mapping = {
                    "image_number": image_counter,
                    "slide_number": current_slide,
                    "position": position.value,
                    "left": None,
                    "top": None,
                    "width": default_width,
                    "height": default_height
                }
                
                ppt_mappings.append(mapping)
                print(f"  Image {image_counter}: {image_name} → Slide {current_slide} ({position.value})")
                image_counter += 1
            else:
                print(f"  ✗ Error: Could not assign position for image {image_counter}")
                image_counter += 1
    
    # Print assignment summary
    _print_assignment_summary(ppt_mappings)
    
    return ppt_mappings

def _find_next_available_slide(position_manager: PositionManager, current_slide: int) -> int:
    """Find next available slide or create new slide number"""
    next_slide = current_slide + 1
    
    # Find next slide that has available positions
    while position_manager.is_slide_full(next_slide):
        next_slide += 1
    
    return next_slide

def _print_assignment_summary(ppt_mappings: List[Dict]):
    """Print summary of position assignments"""
    print("\n=== Assignment Summary ===")
    
    # Group by slide
    slides = {}
    for entry in ppt_mappings:
        slide_num = entry["slide_number"]
        if slide_num not in slides:
            slides[slide_num] = []
        slides[slide_num].append(entry)
    
    for slide_num in sorted(slides.keys()):
        entries = slides[slide_num]
        positions = [entry["position"] for entry in entries]
        images = [entry["image_number"] for entry in entries]
        
        print(f"Slide {slide_num}: Images {images} → Positions {positions}")

def save_to_json(mappings: List[Dict], output_file: str = "mapping.json") -> str:
    """
    Save mappings to JSON file
    
    Args:
        mappings: List of mapping dictionaries
        output_file: Output JSON file path
        
    Returns:
        Path to saved file
    """
    output_path = Path(output_file)
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(mappings, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Saved {len(mappings)} mappings to: {output_path}")
        return str(output_path)
    
    except Exception as e:
        raise IOError(f"Error saving JSON file: {e}")

def validate_mappings(mappings: List[Dict]) -> bool:
    """
    Validate that mappings have required fields
    
    Args:
        mappings: List of mapping dictionaries
        
    Returns:
        True if valid, False otherwise
    """
    required_fields = ['image_number', 'slide_number', 'position']
    
    for i, mapping in enumerate(mappings):
        for field in required_fields:
            if field not in mapping:
                print(f"❌ Mapping {i+1} missing required field: {field}")
                return False
        
        if not isinstance(mapping['image_number'], int) or mapping['image_number'] < 1:
            print(f"❌ Mapping {i+1} has invalid image_number: {mapping['image_number']}")
            return False
        
        if not isinstance(mapping['slide_number'], int) or mapping['slide_number'] < 1:
            print(f"❌ Mapping {i+1} has invalid slide_number: {mapping['slide_number']}")
            return False
    
    print(f"✅ All {len(mappings)} mappings are valid")
    return True

def generate_json_from_docx(docx_path: str, 
                           output_file: str = "mapping.json",
                           default_width: float = 6.0,
                           default_height: float = 4.0,
                           default_position: str = "center") -> str:
    """
    Main function to generate JSON mapping from Word document with collision detection
    
    Args:
        docx_path: Path to Word document containing slide mappings
        output_file: Output JSON file path
        default_width: Default image width in inches
        default_height: Default image height in inches
        default_position: Default position for single images
        
    Returns:
        Path to generated JSON file
    """
    print("=== Generating JSON Mapping from Word Document with Collision Detection ===\n")
    
    # Step 1: Extract slide mappings from docx
    slide_mappings = extract_mappings_from_docx(docx_path)
    
    if not slide_mappings:
        raise ValueError("No slide mappings found in document")
    
    # Step 2: Convert to PPTImageInserter format with collision detection
    ppt_mappings = convert_to_ppt_format_with_collision_detection(
        slide_mappings, 
        default_width=default_width,
        default_height=default_height,
        default_position=default_position
    )
    
    # Step 3: Validate mappings
    if not validate_mappings(ppt_mappings):
        raise ValueError("Generated mappings are invalid")
    
    # Step 4: Save to JSON
    output_path = save_to_json(ppt_mappings, output_file)
    
    print(f"\n🎉 Successfully generated JSON mapping with collision detection!")
    print(f"📄 Source: {docx_path}")
    print(f"📄 Output: {output_path}")
    print(f"📊 Total mappings: {len(ppt_mappings)}")
    
    return output_path

def create_sample_docx():
    """Create a sample Word document with slide mappings"""
    try:
        from docx import Document
        from docx.shared import Inches
        
        doc = Document()
        doc.add_heading('Sample Slide-to-Image Mappings with Collision Detection', 0)
        
        doc.add_paragraph('This document contains slide-to-image mappings for PowerPoint automation.')
        doc.add_paragraph('Format: Slide <number> -- <image1>, <image2>, ...')
        doc.add_paragraph('The script will automatically handle position conflicts using collision detection.')
        doc.add_paragraph()
        
        # Add sample mappings that will trigger collision detection
        sample_mappings = [
            "Slide 1 -- welcome_image.jpg",
            "Slide 2 -- graph1.png, chart1.jpg, diagram1.png, photo1.jpg",  # This will exceed 3 positions
            "Slide 3 -- diagram.png",
            "Slide 4 -- photo1.jpg, photo2.jpg, photo3.png, photo4.jpg, photo5.png",  # This will also exceed
            "Slide 5 -- conclusion.gif"
        ]
        
        for mapping in sample_mappings:
            doc.add_paragraph(mapping)
        
        doc.add_paragraph()
        doc.add_paragraph('Note: Make sure all image files are available in the images directory.')
        doc.add_paragraph('Images exceeding 3 positions per slide will be automatically moved to subsequent slides.')
        
        doc.save('sample_mappings_with_collision_detection.docx')
        print("✅ Created sample document: sample_mappings_with_collision_detection.docx")
        
    except ImportError:
        print("❌ python-docx not installed. Cannot create sample document.")
    except Exception as e:
        print(f"❌ Error creating sample document: {e}")

def main():
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == '--create-sample':
        create_sample_docx()
        return
    
    try:
        # Get inputs
        docx_path = input("Enter Word document path containing slide mappings: ").strip()
        if not docx_path:
            print("❌ Document path is required!")
            return
        
        output_file = input("Enter output JSON file path (press Enter for 'mapping.json'): ").strip()
        if not output_file:
            output_file = "mapping.json"
        
        # Optional parameters
        try:
            width = input("Default image width in inches (press Enter for 6.0): ").strip()
            width = float(width) if width else 6.0
            
            height = input("Default image height in inches (press Enter for 4.0): ").strip()
            height = float(height) if height else 4.0
        except ValueError:
            width, height = 6.0, 4.0
            print("Using default dimensions: 6.0 x 4.0 inches")
        
        position = input("Default position (center/top-left/top-right/bottom-left/bottom-right, press Enter for 'center'): ").strip()
        if position not in ['center', 'top-left', 'top-right', 'bottom-left', 'bottom-right']:
            position = 'center'
        
        # Generate JSON mapping with collision detection
        output_path = generate_json_from_docx(
            docx_path=docx_path,
            output_file=output_file,
            default_width=width,
            default_height=height,
            default_position=position
        )
        
        print(f"\n✅ JSON mapping generated successfully with collision detection!")
        print(f"📁 File saved as: {output_path}")
        print(f"🔧 Collision detection ensures no position conflicts")
        print(f"📊 Excess images automatically moved to subsequent slides")
        
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()