import win32com.client
from typing import Optional, List
from ppt_tools.utils_image import insert_image, extract_images_with_json_metadata
from ppt_tools.utils_slide import (
    get_slide_by_index,
    get_presentation_dimensions,
    get_current_slide_index as _get_current_slide_index
)


class PowerPointController:
    def __init__(self) -> None:
        self.app: Optional[win32com.client.CDispatch] = None

    def connect(self) -> bool:
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            if self.app.Presentations.Count == 0:
                print("❌ No open PowerPoint presentations found.")
                return False
            return True
        except Exception as e:
            print(f"❌ Error connecting to PowerPoint: {e}")
            return False

    def get_current_slide_index(self) -> Optional[int]:
        """Returns the current slide index via the controller."""
        return _get_current_slide_index(self.app)

    def extract_metadata_images(self, slide_index: Optional[int] = None) -> List[tuple[bytes, Optional[str]]]:
        slide = get_slide_by_index(self.app, slide_index)
        if not slide:
            print("⚠️ Could not access slide.")
            return []
        return extract_images_with_json_metadata(slide)

    def insert_metadata_image_offscreen(self, image_path: str, metadata_json: str = "",
                                        apply_style: bool = True) -> None:
        slide = get_slide_by_index(self.app)
        if not slide:
            print("❌ Failed to retrieve current slide.")
            return

        width, height = get_presentation_dimensions(self.app)
        image_width = width * 0.1
        image_height = height * 0.1

        existing_images = sum(
            1 for s in slide.Shapes if s.Type == 13 and s.Left > width
        )

        margin = 10
        left = width + margin
        top = -margin + (existing_images * (image_height + margin))

        insert_image(slide, image_path, left, top, image_width, image_height, metadata_json, apply_style)
        print("🖼️ Corner image added outside slide.")
