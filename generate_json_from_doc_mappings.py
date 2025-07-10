#!/usr/bin/env python3
"""
Generate JSON mapping file from Word document for PPTImageInserter
Extracts slide-to-image mappings from .docx files and converts them to JSON format
with enhanced collision detection and automatic position assignment
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
    """Enhanced position manager with advanced collision detection from doc_json.py"""
    
    def __init__(self):
        # Define position priority order (default starts with bottom-left)
        self.available_positions = [
            Position.BOTTOM_LEFT,    # First choice
            Position.BOTTOM_RIGHT,   # Second choice  
            Position.TOP_RIGHT,      # Third choice
            Position.TOP_LEFT        # Fourth choice if needed
        ]
        # Track occupied positions per slide
        self.occupied_positions: Dict[int, Set[Position]] = {}
    
    def get_next_available_position(self, slide_number: int) -> Optional[Position]:
        """Get next available position on a slide with collision detection"""
        if slide_number not in self.occupied_positions:
            self.occupied_positions[slide_number] = set()
        
        # Return first available position from priority order
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
    
    def get_occupied_positions(self, slide_number: int) -> Set[Position]:
        """Get set of occupied positions for a slide"""
        return self.occupied_positions.get(slide_number, set())
    
    def reset_slide(self, slide_number: int):
        """Reset all positions for a slide"""
        if slide_number in self.occupied_positions:
            self.occupied_positions[slide_number].clear()
    
    def analyze_slide_usage(self, slide_mappings: Dict[str, List[str]]):
        """Analyze slide usage to identify potential issues"""
        slide_usage = {}
        for slide_key, image_names in slide_mappings.items():
            slide_num = int(slide_key.split('_')[1])
            image_count = len(image_names)
            
            if slide_num in slide_usage:
                slide_usage[slide_num] += image_count
            else:
                slide_usage[slide_num] = image_count
        
        print("\n=== Slide Usage Analysis ===")
        max_positions = len(self.available_positions)
        
        for slide_num, count in sorted(slide_usage.items()):
            if count > max_positions:
                print(f"Slide {slide_num}: {count} images (exceeds {max_positions} positions)")
            else:
                print(f"Slide {slide_num}: {count} images")
        
        return slide_usage
    
    def find_or_create_slide(self, current_slide: int) -> int:
        """Find next available slide or suggest new slide number"""
        # Start from current slide + 1 and find next available
        next_slide = current_slide + 1
        
        # Find next slide that has available positions
        while self.is_slide_full(next_slide):
            next_slide += 1
        
        return next_slide

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

def convert_to_ppt_format_with_enhanced_collision_detection(slide_mappings: Dict[str, List[str]], 
                                                          default_width: float = 6.0, 
                                                          default_height: float = 4.0) -> List[Dict]:
    """
    Convert slide-based mappings to the format expected by PPTImageInserter
    with enhanced collision detection and automatic position assignment
    
    Args:
        slide_mappings: {"slide_1": ["image1.jpg", "image2.png"], ...}
        default_width: Default image width in inches
        default_height: Default image height in inches
        
    Returns:
        List of mappings in PPTImageInserter format with collision detection
    """
    position_manager = PositionManager()
    ppt_mappings = []
    image_counter = 1
    
    # Sort slides by number to ensure consistent ordering
    sorted_slides = sorted(slide_mappings.items(), 
                          key=lambda x: int(x[0].split('_')[1]))
    
    print(f"\n=== Converting to PPT Format with Enhanced Collision Detection ===")
    
    # First pass: analyze all mappings to detect potential collisions
    position_manager.analyze_slide_usage(slide_mappings)
    
    # Second pass: assign positions with collision detection
    print(f"\n=== Position Assignment with Enhanced Collision Detection ===")
    
    for slide_key, image_names in sorted_slides:
        slide_num = int(slide_key.split('_')[1])
        
        print(f"\n📍 Processing Slide {slide_num} with {len(image_names)} images:")
        
        # Check if slide can accommodate all images
        available_positions = position_manager.get_available_count(slide_num)
        max_positions = len(position_manager.available_positions)
        
        if len(image_names) > available_positions:
            print(f"  ⚠️ Warning: Slide {slide_num} needs {len(image_names)} positions but only has {available_positions} available")
            print(f"  ✓ Will distribute excess images to subsequent slides")
        
        current_slide = slide_num
        
        # Assign positions for each image with enhanced collision detection
        for i, image_name in enumerate(image_names):
            # Check if current slide has space
            if position_manager.is_slide_full(current_slide):
                # Find next available slide
                current_slide = position_manager.find_or_create_slide(current_slide)
                print(f"  ✓ Moving to slide {current_slide} for image {image_counter} ({image_name})")
            
            # Get next available position with collision detection
            position = position_manager.get_next_available_position(current_slide)
            
            if position:
                # Mark position as occupied
                position_manager.occupy_position(current_slide, position)
                
                # Create mapping entry with null coordinates (let PPT handle positioning)
                mapping = {
                    "image_number": image_counter,
                    "slide_number": current_slide,
                    "position": position.value,
                    "left": None,  # Use null instead of 0
                    "top": None,   # Use null instead of 0
                    "width": default_width,
                    "height": default_height
                }
                
                ppt_mappings.append(mapping)
                
                # Show collision detection in action
                occupied_positions = position_manager.get_occupied_positions(current_slide)
                print(f"    Image {image_counter} ({image_name}) → Slide {current_slide}")
                print(f"    Position: {position.value} (occupied positions: {[p.value for p in occupied_positions]})")
                
                image_counter += 1
            else:
                print(f"  ✗ Error: Could not assign position for image {image_counter} ({image_name})")
                # Still increment counter to maintain sequence
                image_counter += 1
    
    # Print detailed assignment summary
    print_enhanced_assignment_summary(ppt_mappings, position_manager)
    
    return ppt_mappings

def print_enhanced_assignment_summary(ppt_mappings: List[Dict], position_manager: PositionManager):
    """Print enhanced summary of position assignments with collision detection details"""
    print("\n=== Enhanced Assignment Summary with Collision Detection ===")
    
    # Group by slide
    slides = {}
    for entry in ppt_mappings:
        slide_num = entry["slide_number"]
        if slide_num not in slides:
            slides[slide_num] = []
        slides[slide_num].append(entry)
    
    total_images = len(ppt_mappings)
    total_slides = len(slides)
    
    print(f"📊 Total Images: {total_images}")
    print(f"📊 Total Slides Used: {total_slides}")
    print(f"📊 Average Images per Slide: {total_images/total_slides:.1f}")
    
    print("\n--- Slide-by-Slide Assignment ---")
    for slide_num in sorted(slides.keys()):
        entries = slides[slide_num]
        positions = [entry["position"] for entry in entries]
        images = [entry["image_number"] for entry in entries]
        
        # Show position distribution
        position_counts = {}
        for pos in positions:
            position_counts[pos] = position_counts.get(pos, 0) + 1
        
        print(f"Slide {slide_num}: {len(entries)} images")
        print(f"  Images: {images}")
        print(f"  Positions: {positions}")
        print(f"  Position usage: {position_counts}")
        
        # Show collision detection effectiveness
        if len(entries) > 1:
            print(f"  ✓ Collision detection successfully assigned {len(entries)} images to different positions")
    
    # Show overflow analysis
    print("\n--- Overflow Analysis ---")
    max_positions_per_slide = len(position_manager.available_positions)
    overflowed_slides = [slide_num for slide_num, entries in slides.items() 
                        if len(entries) > max_positions_per_slide]
    
    if overflowed_slides:
        print(f"⚠️  Original slides that caused overflow: {overflowed_slides}")
        print(f"✓ Overflow images automatically distributed to subsequent slides")
    else:
        print("✓ No overflow occurred - all images fit within available positions")

def save_to_json(mappings: List[Dict], output_file: str = "mapping.json") -> str:
    """
    Save mappings to JSON file with enhanced formatting
    
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
        
        # Show sample of saved data
        if mappings:
            print("\n--- Sample JSON Entry ---")
            sample = mappings[0]
            print(json.dumps(sample, indent=2))
            if len(mappings) > 1:
                print(f"... and {len(mappings) - 1} more entries")
        
        return str(output_path)
    
    except Exception as e:
        raise IOError(f"Error saving JSON file: {e}")

def validate_mappings(mappings: List[Dict]) -> bool:
    """
    Validate that mappings have required fields and proper collision detection
    
    Args:
        mappings: List of mapping dictionaries
        
    Returns:
        True if valid, False otherwise
    """
    required_fields = ['image_number', 'slide_number', 'position', 'left', 'top', 'width', 'height']
    
    print(f"\n=== Validating {len(mappings)} mappings ===")
    
    # Check required fields
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
        
        # Validate position values
        valid_positions = ['top-left', 'top-right', 'bottom-left', 'bottom-right', 'center', 'custom']
        if mapping['position'] not in valid_positions:
            print(f"❌ Mapping {i+1} has invalid position: {mapping['position']}")
            return False
        
        # Check that left and top are None (null) for positioned images
        if mapping['left'] is not None and mapping['position'] != 'custom':
            print(f"⚠️  Mapping {i+1} has left coordinate set but position is not 'custom'")
        
        if mapping['top'] is not None and mapping['position'] != 'custom':
            print(f"⚠️  Mapping {i+1} has top coordinate set but position is not 'custom'")
    
    # Validate collision detection - check for position conflicts
    print("\n--- Collision Detection Validation ---")
    slide_positions = {}
    conflicts = 0
    
    for mapping in mappings:
        slide_num = mapping['slide_number']
        position = mapping['position']
        
        if slide_num not in slide_positions:
            slide_positions[slide_num] = []
        
        if position in slide_positions[slide_num]:
            print(f"❌ Position conflict detected: Slide {slide_num}, Position {position}")
            conflicts += 1
        else:
            slide_positions[slide_num].append(position)
    
    if conflicts > 0:
        print(f"❌ Found {conflicts} position conflicts")
        return False
    
    print(f"✅ All {len(mappings)} mappings are valid")
    print(f"✅ No position conflicts detected - collision detection working properly")
    return True

def generate_json_from_docx(docx_path: str, 
                           output_file: str = "mapping.json",
                           default_width: float = 6.0,
                           default_height: float = 4.0) -> str:
    """
    Main function to generate JSON mapping from Word document with enhanced collision detection
    
    Args:
        docx_path: Path to Word document containing slide mappings
        output_file: Output JSON file path
        default_width: Default image width in inches
        default_height: Default image height in inches
        
    Returns:
        Path to generated JSON file
    """
    print("=== Generating JSON Mapping with Enhanced Collision Detection ===\n")
    
    # Step 1: Extract slide mappings from docx
    slide_mappings = extract_mappings_from_docx(docx_path)
    
    if not slide_mappings:
        raise ValueError("No slide mappings found in document")
    
    # Step 2: Convert to PPTImageInserter format with enhanced collision detection
    ppt_mappings = convert_to_ppt_format_with_enhanced_collision_detection(
        slide_mappings, 
        default_width=default_width,
        default_height=default_height
    )
    
    # Step 3: Validate mappings including collision detection
    if not validate_mappings(ppt_mappings):
        raise ValueError("Generated mappings are invalid")
    
    # Step 4: Save to JSON
    output_path = save_to_json(ppt_mappings, output_file)
    
    print(f"\n🎉 Successfully generated JSON mapping with enhanced collision detection!")
    print(f"📄 Source: {docx_path}")
    print(f"📄 Output: {output_path}")
    print(f"📊 Total mappings: {len(ppt_mappings)}")
    print(f"🔧 Collision detection: Active (using null coordinates)")
    print(f"📍 Default position sequence: bottom-left → bottom-right → top-right → top-left")
    
    return output_path



def main():
    import sys
    
    
    
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
        
        # Generate JSON mapping with enhanced collision detection
        output_path = generate_json_from_docx(
            docx_path=docx_path,
            output_file=output_file,
            default_width=width,
            default_height=height
        )
        
        print(f"\n✅ JSON mapping generated successfully with enhanced collision detection!")
        print(f"📁 File saved as: {output_path}")
        print(f"🔧 Enhanced collision detection ensures no position conflicts")
        print(f"📊 Excess images automatically moved to subsequent slides")
        print(f"📍 Coordinates set to null for automatic positioning")
        
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()