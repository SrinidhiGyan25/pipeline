import os
import json
import re
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from pathlib import Path
import logging
from script import CanvasConverter
from generate_json_from_doc_mappings import extract_mappings_from_docx, save_to_json
from image_auto import PPTImageInserter
from pathlib import Path
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import re
from docx import Document
# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

@dataclass
class ContentMetadata:
    """Data class to store content metadata"""
    document_name: str
    gpt_canvas: bool
    gpt_canvas_without_speaker_notes: bool
    ppt_needs_images: bool
    ppt_has_images: bool
    ppt_with_speaker_notes: bool
    gpt_canvas_link: Optional[str] = None
    ppt_drive_link: Optional[str] = None
    image_folder_link: Optional[str] = None
    image_mapping_file: Optional[str] = None
    image_mappings: Optional[Dict[str, List[str]]] = None


class ContentProcessor:
    """Main driver class for content processing automation"""
    
    def __init__(self, output_dir: str = "output"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        # Initialize component modules (these would be your actual scripts)
        self.canvas_to_ppt_converter = CanvasToPPTConverter()
        self.image_inserter = ImageInserter()
        self.json_mapping_generator = JSONMappingGenerator()
        self.drive_handler = DriveHandler()
        
    def parse_metadata(self, content: str) -> ContentMetadata:
        """Parse metadata from the document content"""
        logger.info("Parsing metadata from content...")
        
        # Extract document name
        doc_name_match = re.search(r'\*\*Document Name\*\*\s*:\s*(.+)', content)
        document_name = doc_name_match.group(1).strip() if doc_name_match else "unknown"
        
        # Extract Y/N flags
        def extract_flag(pattern: str) -> bool:
            match = re.search(pattern, content, re.IGNORECASE)
            return match.group(1).strip().upper() == 'Y' if match else False
        
        gpt_canvas = extract_flag(r'\*\*Gpt canvas\*\*.*?([YN])')
        gpt_canvas_without_notes = extract_flag(r'\*\*Gpt canvas without.*?speaker notes\*\*.*?([YN])')
        ppt_needs_images = extract_flag(r'\*\*Ppt needs images\*\*.*?([YN])')
        ppt_has_images = extract_flag(r'\*\*PPT has images\*\*.*?([YN])')
        ppt_with_speaker_notes = extract_flag(r'\*\*Ppt with speaker notes\*\*.*?([YN])')
        
        # Extract links
        canvas_link_match = re.search(r'https://chatgpt\.com/canvas/shared/[a-zA-Z0-9]+', content)
        canvas_link = canvas_link_match.group(0) if canvas_link_match else None
        
        ppt_link_match = re.search(r'https://drive\.google\.com/.*?(?=\s|$)', content)
        ppt_link = ppt_link_match.group(0) if ppt_link_match else None
        
        image_folder_match = re.search(r'(https://drive\.google\.com/drive/folders/[a-zA-Z0-9_-]+)', content)
        image_folder_link = image_folder_match.group(1) if image_folder_match else None
        
        # Extract image mappings
        image_mappings = self._extract_image_mappings(content)
        
        return ContentMetadata(
            document_name=document_name,
            gpt_canvas=gpt_canvas,
            gpt_canvas_without_speaker_notes=gpt_canvas_without_notes,
            ppt_needs_images=ppt_needs_images,
            ppt_has_images=ppt_has_images,
            ppt_with_speaker_notes=ppt_with_speaker_notes,
            gpt_canvas_link=canvas_link,
            ppt_drive_link=ppt_link,
            image_folder_link=image_folder_link,
            image_mappings=image_mappings
        )
    
    def _extract_image_mappings(self, content: str) -> Dict[str, List[str]]:
        """Extract image mappings from content"""
        mappings = {}
        
        # Find all slide to image mappings
        slide_pattern = r'Slide\s+(\d+)\s*--\s*(.+)'
        matches = re.findall(slide_pattern, content)
        
        for slide_num, images_str in matches:
            # Split images by comma and clean up
            images = [img.strip() for img in images_str.split(',')]
            mappings[f"slide_{slide_num}"] = images
        
        return mappings
    
    def process_content(self, content: str) -> Dict[str, str]:
        """Main processing pipeline"""
        logger.info("Starting content processing pipeline...")
        
        # Parse metadata
        metadata = self.parse_metadata(content)
        logger.info(f"Processing: {metadata.document_name}")
        
        # Create processing results dictionary
        results = {
            "document_name": metadata.document_name,
            "processing_steps": []
        }
        
        # Determine processing path based on flags
        if metadata.gpt_canvas:
            canvas_results = self._process_canvas_path(metadata)
            results["processing_steps"].extend(canvas_results.get("processing_steps", []))
            results.update({k: v for k, v in canvas_results.items() if k != "processing_steps"})

        elif metadata.ppt_drive_link:
            ppt_results = self._process_ppt_path(metadata)
            results["processing_steps"].extend(ppt_results.get("processing_steps", []))
            results.update({k: v for k, v in ppt_results.items() if k != "processing_steps"})

        else:
            logger.warning("No valid input source found (Canvas or PPT)")
            results["error"] = "No valid input source"
            return results
        
        # Generate final JSON mapping if needed
        mapping_path = f"output/{metadata.document_name}_mapping.json"
        if metadata.ppt_needs_images and not Path(mapping_path).exists():
            self._generate_json_mapping(metadata, results)

        
        logger.info("Content processing completed successfully")
        return results
    
    def _process_canvas_path(self, metadata: ContentMetadata) -> Dict[str, str]:
        """Process content when Canvas is the source"""
        logger.info("Processing Canvas path...")
        results = {
            "processing_steps": []
        }
        
        # Step 1: Download Canvas content
        if metadata.gpt_canvas_link:
            canvas_content = self.drive_handler.download_canvas_content(metadata.gpt_canvas_link)
            results["canvas_downloaded"] = "success"
            results["processing_steps"].append("Canvas content downloaded")
        else:
            logger.error("Canvas link not found")
            results["error"] = "Canvas link not found"
            return results
        
        # Step 2: Convert Canvas to PPT
        ppt_path = self.canvas_to_ppt_converter.convert(
        canvas_link=metadata.gpt_canvas_link,
        include_speaker_notes=not metadata.gpt_canvas_without_speaker_notes
    )
        results["ppt_generated"] = str(ppt_path)
        results["processing_steps"].append("Canvas converted to PPT")
        
        # Step 3: Handle images if needed
        if metadata.ppt_needs_images and metadata.image_folder_link:
            results.update(self._process_images(metadata, ppt_path))
        
        return results
    
    def _process_ppt_path(self, metadata: ContentMetadata) -> Dict[str, str]:
        """Process content when PPT is the source"""
        logger.info("Processing PPT path...")
        results = {
            "processing_steps": []
        }
        
        # Step 1: Download PPT
        ppt_path = self.drive_handler.download_ppt(metadata.ppt_drive_link)
        results["ppt_downloaded"] = str(ppt_path)
        results["processing_steps"].append("PPT downloaded")
        
        # Step 2: Handle images if needed
        if metadata.ppt_needs_images and not metadata.ppt_has_images:
            if metadata.image_folder_link:
                results.update(self._process_images(metadata, ppt_path))
            else:
                logger.warning("PPT needs images but no image folder link provided")
        
        return results
    
    def _process_images(self, metadata: ContentMetadata, ppt_path: Path) -> Dict[str, str]:
        """Process and insert images into PPT"""
        logger.info("Processing images...")
        results = {"processing_steps": []}  # ✅ Fixed: Initialize processing_steps
        
        # Generate JSON mapping FIRST (before downloading images)
        mapping_file = f"output/{metadata.document_name}_mapping.json"
        if not Path(mapping_file).exists():
            self._generate_json_mapping(metadata, results)
        
        # ✅ Ensure metadata.image_mapping_file is correctly set
        if not metadata.image_mapping_file:
            metadata.image_mapping_file = mapping_file
        
        
        # Download images from Drive to output directory
        image_paths = self.drive_handler.download_images(metadata.image_folder_link)
        results["images_downloaded"] = len(image_paths)
        results["processing_steps"].append(f"Downloaded {len(image_paths)} images")
        
        # ✅ Fixed: Pass the correct image directory (where images were actually downloaded)
        if metadata.image_mappings and image_paths:
            updated_ppt_path = self.image_inserter.insert_images(
                ppt_path=ppt_path, 
                image_paths=image_paths, 
                mappings=metadata.image_mappings,
                mapping_file=metadata.image_mapping_file,
                image_dir=str(self.output_dir)  # ✅ Pass actual image directory
            )
            results["images_inserted"] = str(updated_ppt_path)
            results["processing_steps"].append("Images inserted into PPT")
        
        return results
    
    def _generate_json_mapping(self, metadata: ContentMetadata, results: Dict[str, str]):
        """Generate JSON mapping file"""
        logger.info("Generating JSON mapping...")
        
        # Generate and save the mapping JSON
        output_file = f"output/{metadata.document_name}_mapping.json"
        json_path = self.json_mapping_generator.generate(metadata, output_file=output_file)
        results["json_mapping_generated"] = str(json_path)
        metadata.image_mapping_file = str(json_path)
        
        # ✅ Load mappings from the file into metadata
        with open(json_path, "r") as f:
            loaded_mappings = json.load(f)
            # ✅ Convert to the format expected by the rest of the code
            metadata.image_mappings = loaded_mappings
        
        metadata.image_mapping_file = json_path


# Component classes (these would be your actual implementation modules)

class CanvasToPPTConverter:
    """Wrapper to use the CanvasConverter from script.py"""

    def __init__(self):
        self.converter = CanvasConverter()

    def convert(self, canvas_link: str, include_speaker_notes: bool = True) -> Path:
        # Output dir can be fixed or dynamic
        ppt_path = self.converter.convert(
            url=canvas_link,
            output_dir="output"
        )
        return ppt_path

class ImageInserter:
    """Wrapper for the PPTImageInserter from image_auto.py"""

    def __init__(self):
        self.inserter = PPTImageInserter()

    def insert_images(self, ppt_path: Path, image_paths: List[Path], mappings: Dict[str, List[str]], 
                     mapping_file: Optional[str] = None, image_dir: str = None) -> Path:
        """
        ✅ Fixed: Updated to handle correct image directory and mapping format
        """
        # Use provided image_dir or default to ppt directory
        if image_dir is None:
            image_dir = str(ppt_path.parent)
        
        # Use provided mapping file or default
        if mapping_file is None:
            mapping_file = str(ppt_path.parent / "mapping.json")
        
        # Generate output filename
        output_file = str(ppt_path.parent / f"with_images_{ppt_path.name}")
        
        print(f"📌 Starting image insertion using: {mapping_file}")
        print(f"📌 Image directory: {image_dir}")
        print(f"📌 PPT file: {ppt_path}")
        print(f"📌 Output file: {output_file}")
        
        # ✅ Verify files exist before processing
        if not os.path.exists(mapping_file):
            raise FileNotFoundError(f"Mapping file not found: {mapping_file}")
        
        if not os.path.exists(str(ppt_path)):
            raise FileNotFoundError(f"PPT file not found: {ppt_path}")
        
        # Call the actual image insertion
        success = self.inserter.process_mappings(
            ppt_file=str(ppt_path),
            image_dir=image_dir,
            mapping_file=mapping_file,
            output_file=output_file
        )
        
        if not success:
            raise RuntimeError("Image insertion failed")
        
        return Path(output_file)

class JSONMappingGenerator:
    """Generates JSON mapping file for image insertion"""

    def generate(self, metadata, output_file="mapping.json") -> Path:
        """Generate JSON mapping based on text slide-to-image mapping in the .docx"""

        # Use document name to find the docx
        docx_path = f"{metadata.document_name}.docx"
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"Document file not found: {docx_path}")

        print(f"📄 Generating JSON mapping from: {docx_path}")
        mappings = extract_mappings_from_docx(docx_path)
        
        # ✅ Convert slide-based mappings to image-based mappings for PPTImageInserter
        image_mappings = self._convert_to_image_mappings(mappings)
        
        # Save in the format expected by PPTImageInserter
        save_to_json(image_mappings, output_file)
        return Path(output_file)
    
    def _convert_to_image_mappings(self, slide_mappings: Dict[str, List[str]]) -> List[Dict]:
        """
        ✅ Convert slide-based mappings to image-based mappings expected by PPTImageInserter
        
        Input format (from docx): {"slide_1": ["image1.jpg", "image2.jpg"], "slide_2": ["image3.jpg"]}
        Output format (for PPTImageInserter): [{"image_number": 1, "slide_number": 1, "position": "center"}, ...]
        """
        image_mappings = []
        image_counter = 1
        
        for slide_key, image_names in slide_mappings.items():
            # Extract slide number from key like "slide_1"
            slide_num = int(slide_key.split('_')[1])
            
            for image_name in image_names:
                mapping = {
                    "image_number": image_counter,
                    "slide_number": slide_num,
                    "position": "center",  # Default position
                    "left": None,
                    "top": None,
                    "width": 6.0,  # Default width in inches
                    "height": 4.0  # Default height in inches
                }
                image_mappings.append(mapping)
                image_counter += 1
        
        return image_mappings

class DriveHandler:
    """Handles Google Drive operations using PyDrive"""

    def __init__(self):
        self.gauth = GoogleAuth()
        self.gauth.LocalWebserverAuth()  # Will open browser on first run
        self.drive = GoogleDrive(self.gauth)

    def extract_file_id(self, url: str) -> str:
        # Try /d/<id> (file link)
        match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
        if not match:
            # Try id=<id> (alternate file link format)
            match = re.search(r'id=([a-zA-Z0-9_-]+)', url)
        if not match:
            # Try /folders/<id> (folder link)
            match = re.search(r'/folders/([a-zA-Z0-9_-]+)', url)
        return match.group(1) if match else None

    def download_ppt(self, drive_link: str) -> Path:
        file_id = self.extract_file_id(drive_link)
        if not file_id:
            raise ValueError("Invalid Drive PPT link format")

        file = self.drive.CreateFile({'id': file_id})
        file_name = file['title']
        out_path = Path("output") / file_name
        file.GetContentFile(str(out_path))
        logger.info(f"📥 Downloaded PPT to: {out_path}")
        return out_path

    def download_images(self, folder_link: str) -> List[Path]:
        folder_id = self.extract_file_id(folder_link)
        # if not file_id:
        #     raise ValueError("Invalid Drive folder link format")

        file_list = self.drive.ListFile({
            'q': f"'{folder_id}' in parents and trashed=false"
        }).GetList()

        output_dir = Path("output")
        image_paths = []

        for file in file_list:
            if file['mimeType'].startswith('image/'):
                file_name = file['title']
                out_path = output_dir / file_name
                file.GetContentFile(str(out_path))
                logger.info(f"📥 Downloaded image: {file_name}")
                image_paths.append(out_path)

        logger.info(f"📥 Total images downloaded: {len(image_paths)}")
        return image_paths
    
    def download_canvas_content(self, canvas_link: str) -> str:
        logger.info(f"Skipping download: Canvas content will be handled by CanvasConverter.")
        return "Canvas content is handled separately"

# Usage example
def read_docx_text(path: str) -> str:
    doc = Document(path)
    return "\n".join([p.text for p in doc.paragraphs])

def main():
    import sys
    if len(sys.argv) < 2:
        print("❌ Please provide path to metadata .docx file.\nUsage: python driver.py <metadata_file.docx>")
        return
    
    metadata_path = sys.argv[1]
    
    try:
        content_text = read_docx_text(metadata_path)
        processor = ContentProcessor()
        results = processor.process_content(content_text)
        print("\n✅ Processing complete.\n")
        print(json.dumps(results, indent=2))
    except Exception as e:
        print(f"❌ Error during processing: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()