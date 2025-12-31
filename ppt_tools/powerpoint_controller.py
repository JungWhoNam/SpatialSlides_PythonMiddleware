import win32com.client
from typing import Optional, List
import logging
from ppt_tools import utils_image
from ppt_tools import utils_slide


class PowerPointController:
    def __init__(self) -> None:
        self.app: Optional[win32com.client.CDispatch] = None

    def connect(self) -> bool:
        """Connects to an active PowerPoint application."""
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            if self.app.Presentations.Count == 0:
                logging.error("No open PowerPoint presentations found.")
                return False
            return True
        except Exception as e:
            logging.error(f"Error connecting to PowerPoint via COM: {e}")
            self.app = None
            return False

    def get_current_slide_index(self) -> Optional[int]:
        """Returns the current slide index."""
        return utils_slide.get_current_slide_index(self.app)

    def get_current_click_index(self) -> Optional[int]:
        """Returns the current click index."""
        return utils_slide.get_current_click_index(self.app)

    def is_presenter_mode(self) -> Optional[bool]:
        """
        Returns True if PowerPoint is in presenter mode,
        False if in edit mode,
        or None if the state could not be determined.
        """
        return utils_slide.is_presenter_mode(self.app)

    def extract_metadata_images(self, slide_index: Optional[int] = None) -> List[tuple[bytes, Optional[str]]]:
        """Extracts images and metadata from the given or current slide."""
        slide = utils_slide.get_slide_by_index(self.app, slide_index)
        if not slide:
            logging.warning("Could not access slide for image extraction.")
            return []
        return utils_image.extract_images_with_json_metadata(slide)

    def extract_metadata_only(self, slide_index: Optional[int] = None) -> List[str]:
        """Extracts only the JSON metadata strings from the given or current slide."""
        slide = utils_slide.get_slide_by_index(self.app, slide_index)
        if not slide:
            logging.warning("Could not access slide for metadata extraction.")
            return []
        return utils_image.extract_json_metadata_only(slide)

    def extract_all_metadata_images(self) -> List[tuple[bytes, Optional[str]]]:
        """Extracts all (image, metadata) pairs from all slides."""
        result: List[tuple[bytes, Optional[str]]] = []

        pres = utils_slide.get_active_presentation(self.app)
        if not pres:
            logging.warning("Cannot extract all metadata: No active presentation.")
            return []

        for slide in pres.Slides:
            result.extend(utils_image.extract_images_with_json_metadata(slide))

        return result

    def insert_metadata_image_offscreen(
            self,
            image_path: str,
            metadata_json: str = "",
            apply_style: bool = True
    ) -> None:
        """
        Inserts an image outside the visible slide area (e.g., for embedding metadata).
        Assumes the image is square for sizing purposes.
        """
        slide = utils_slide.get_slide_by_index(self.app)
        if not slide:
            logging.error("Failed to retrieve current slide for metadata insertion.")
            return

        # Define fixed height, and infer width using square assumption
        width, height = utils_slide.get_presentation_dimensions(self.app)
        image_height = height * 0.1
        image_width = image_height * 1.0

        # Stack off-screen images vertically
        margin = 10
        left = width + margin
        existing_images = sum(
            1 for shape in slide.Shapes if shape.Type == 13 and shape.Left > width
        )
        top = -margin + (existing_images * (image_height + margin))

        utils_image.insert_image(
            slide,
            image_path,
            left,
            top,
            image_width,
            image_height,
            metadata_json,
            apply_style
        )
        logging.info("Metadata image added offscreen.")