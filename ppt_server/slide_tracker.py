from typing import Optional
from ppt_tools.powerpoint_controller import PowerPointController


class SlideTracker:
    def __init__(self, ppt_controller: PowerPointController):
        self.ppt_controller = ppt_controller
        self.last_slide_index: Optional[int] = None

    def check_slide_change(self) -> Optional[int]:
        """Returns the current slide index if it has changed since the last check."""
        current_slide_index = self.ppt_controller.get_current_slide_index()
        if current_slide_index is None or current_slide_index == self.last_slide_index:
            return None

        self.last_slide_index = current_slide_index
        print(f"\n🔄 Slide changed to: {current_slide_index}")
        return current_slide_index
