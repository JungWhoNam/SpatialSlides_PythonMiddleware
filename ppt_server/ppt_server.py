import sys
import time
import json
from typing import Optional, List, Tuple, Literal
import logging
from .zmq_handler import ZMQHandler
from ppt_tools.powerpoint_controller import PowerPointController
from .slide_tracker import SlideTracker
from .client_message_handler import handle_message
from .server_action import ServerAction


class PowerPointServer:
    """Main ppt_server that integrates PowerPoint tracking and ZMQ communication."""

    def __init__(self, interval=1.0):
        """
        Initializes PowerPoint tracking and ZMQ messaging.

        Args:
            interval (float): Polling interval for monitoring PowerPoint slides.
        """
        self.ppt_controller = PowerPointController()
        self.zmq_handler = ZMQHandler()
        self.slide_tracker = SlideTracker(self.ppt_controller)
        self.running = True  # Flag to control ppt_server loop
        self.interval = interval  # Polling interval

    def start(self):
        """Starts tracking PowerPoint slides and processing client messages."""
        if not self.ppt_controller.connect():
            logging.error("Could not connect to PowerPoint. Exiting.")
            return

        logging.info("PowerPoint Server is running. Processing incoming messages...")

        try:
            while self.running:
                start_time = time.time()

                # Process incoming messages from Unity
                requested_actions: set[ServerAction] = set()

                def wrapper(msg_parts):
                    result = handle_message(msg_parts, self.ppt_controller)
                    if result in (
                            ServerAction.SEND_CURRENT_VIEWS, ServerAction.SEND_ALL_VIEWS,
                            ServerAction.SEND_CURRENT_MODE):
                        requested_actions.add(result)

                self.zmq_handler.process_queue(wrapper)

                # 1. Prioritize checking for a slide change.
                slide_index = self.slide_tracker.check_slide_change()
                if slide_index is not None:
                    # A slide change occurred, send a single update for the new slide.
                    metadata_list = self.ppt_controller.extract_metadata_only(slide_index)
                    parts = self.build_metadata_only_message("CurrentViewRefs", metadata_list, slide_index)
                    self.zmq_handler.send_multipart(parts)

                # 2. Only if the slide has NOT changed, check for an animation change.
                else:
                    animation_state = self.slide_tracker.check_animation_change()
                    if animation_state:
                        slide_index, animation_step = animation_state
                        # No images are sent.
                        parts = self.build_animation_step_message(slide_index, animation_step)
                        self.zmq_handler.send_multipart(parts)

                # Handle explicit requests from the client
                if ServerAction.SEND_CURRENT_VIEWS in requested_actions:
                    target_index = self.ppt_controller.get_current_slide_index()
                    # images = self.ppt_controller.extract_metadata_images(target_index)
                    # parts = self.build_view_message("CurrentViews", images, target_index)
                    metadata_list = self.ppt_controller.extract_metadata_only(target_index)
                    parts = self.build_metadata_only_message("CurrentViewRefs", metadata_list, target_index)
                    self.zmq_handler.send_multipart(parts)

                if ServerAction.SEND_ALL_VIEWS in requested_actions:
                    images = self.ppt_controller.extract_all_metadata_images()
                    parts = self.build_view_message("AllViews", images)
                    self.zmq_handler.send_multipart(parts)

                # Check for mode changes (edit <-> present) or explicit request
                mode = self.slide_tracker.check_mode_change()
                if mode or ServerAction.SEND_CURRENT_MODE in requested_actions:
                    target_mode = mode or ("present" if self.ppt_controller.is_presenter_mode() else "edit")
                    parts = self.build_current_mode_message(target_mode)
                    self.zmq_handler.send_multipart(parts)

                time.sleep(max(0, self.interval - (time.time() - start_time)))

        except KeyboardInterrupt:
            self.shutdown()

    def build_view_message(self, header: str, images: List[Tuple[bytes, Optional[str]]],
                           slide_index: Optional[int] = None) -> List[bytes]:
        """
        Builds a multipart message containing image + metadata pairs, with optional slide index metadata.

        Args:
            header (str): The message type header (e.g., "SlideChanged", "AllViews")
            images (List): List of (image_bytes, metadata_json) tuples.
            slide_index (Optional[int]): If provided, will be included in the second part as JSON.

        Returns:
            List[bytes]: Multipart message parts
        """
        parts = [header.encode("utf-8")]

        if slide_index is not None:
            parts.append(json.dumps({"slide": slide_index}).encode("utf-8"))

        for image_bytes, alt_text in images:
            meta = alt_text or "{}"
            parts.append(meta.encode("utf-8"))
            parts.append(image_bytes)

        return parts

    def build_metadata_only_message(self, header: str, metadata_list: List[str], slide_index: int) -> List[bytes]:
        """Builds a multipart message containing only metadata references."""
        parts = [header.encode("utf-8")]
        parts.append(json.dumps({"slide": slide_index}).encode("utf-8"))

        for meta_json in metadata_list:
            parts.append(meta_json.encode("utf-8"))

        return parts

    def build_current_mode_message(self, mode: Literal["edit", "present"]) -> List[bytes]:
        """
        Builds a multipart message for mode change notification.
        """
        return [
            b"CurrentMode",
            json.dumps({"mode": mode}).encode("utf-8")
        ]

    def build_animation_step_message(self, slide_index: int, animation_step: int) -> List[bytes]:
        """
        Builds a lightweight message to notify the client of an animation step change.
        """
        parts = [
            b"AnimationStep",
            json.dumps({
                "slide": slide_index,
                "animation_step": animation_step
            }).encode("utf-8")
        ]
        return parts

    def shutdown(self):
        """Shuts down the ppt_server."""
        self.running = False
        self.zmq_handler.close()
        logging.info("PowerPoint Server shut down successfully.")
        sys.exit(0)
