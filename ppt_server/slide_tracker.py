from typing import Optional, Tuple
from ppt.ppt_controller import PowerPointController


class SlideTracker:
    """Tracks slide changes in PowerPoint."""

    def __init__(self, ppt_controller: PowerPointController):
        """
        Initializes the slide tracker.

        Args:
            ppt_controller (PowerPointController): PowerPoint interaction instance.
        """
        self.ppt_controller = ppt_controller
        self.last_slide_index: Optional[int] = None

    def check_slide_change(self) -> Optional[Tuple[int, Optional[list]]]:
        """Checks if the slide has changed and returns slide index & notes if changed."""
        current_slide_index = self.ppt_controller.get_current_slide_index()

        if current_slide_index is None or current_slide_index == self.last_slide_index:
            return None

        self.last_slide_index = current_slide_index
        print(f"\n🔄 Slide changed to: {current_slide_index}")

        alt_texts = self.ppt_controller.get_image_alt_text_from_slide(current_slide_index)
        return current_slide_index, alt_texts
