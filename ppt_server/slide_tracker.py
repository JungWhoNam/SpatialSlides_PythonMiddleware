from typing import Optional, Literal
from ppt_tools.powerpoint_controller import PowerPointController


class SlideTracker:
    def __init__(self, ppt_controller: PowerPointController):
        self.ppt_controller = ppt_controller
        self.last_slide_index: Optional[int] = None
        self.last_mode: Optional[Literal["edit", "present"]] = None
        self.last_animation_step: Optional[int] = None

    def check_slide_change(self) -> Optional[int]:
        """Returns the current slide index if it has changed since the last check."""
        current_slide_index = self.ppt_controller.get_current_slide_index()
        if current_slide_index is None or current_slide_index == self.last_slide_index:
            return None

        self.last_slide_index = current_slide_index
        print(f"\n🔄 Slide changed to: {current_slide_index}")
        return current_slide_index

    def check_mode_change(self) -> Optional[Literal["edit", "present"]]:
        """
        Checks whether the presentation mode has changed.
        Returns the new mode string if it changed: "edit" or "present".
        Returns None if no change or error.
        """
        is_presenting = self.ppt_controller.is_presenter_mode()
        if is_presenting is None:
            return None  # Unable to determine

        current_mode: Literal["edit", "present"] = "present" if is_presenting else "edit"

        if current_mode != self.last_mode:
            print(f"🎬 Mode changed to: {current_mode}")
            self.last_mode = current_mode
            return current_mode

        return None

    def check_animation_change(self) -> Optional[tuple[int, int]]:
        """
        Checks if the animation step on the current slide has changed.
        This method is designed to work in tandem with check_slide_change.
        """
        if not self.ppt_controller.is_presenter_mode():
            if self.last_animation_step is not None:
                self.last_animation_step = None
            return None

        try:
            current_slide_index = self.ppt_controller.get_current_slide_index()
            current_animation_step = self.ppt_controller.get_current_click_index()

            if current_slide_index is None or current_animation_step is None:
                return None

            if current_slide_index != self.last_slide_index:
                self.last_animation_step = 0
                return None

            # Now, if we're on the same slide, check if only the animation step has changed.
            if current_animation_step != self.last_animation_step:
                self.last_animation_step = current_animation_step
                print(f"\n🔄 Animation changed to:  {current_animation_step}")
                return current_slide_index, current_animation_step

        except Exception:
            return None

        return None
