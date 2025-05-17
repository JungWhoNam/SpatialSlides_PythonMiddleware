import sys
import time
import json
from typing import Optional, List, Tuple
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
            print("❌ Could not connect to PowerPoint. Exiting.")
            return

        print("✅ PowerPoint Server is running. Processing incoming messages...")

        try:
            while self.running:
                start_time = time.time()

                # Process incoming messages from Unity
                requested_actions: set[ServerAction] = set()

                def wrapper(msg_parts):
                    result = handle_message(msg_parts, self.ppt_controller)
                    if result in (ServerAction.SEND_CURRENT_VIEWS, ServerAction.SEND_ALL_VIEWS):
                        requested_actions.add(result)

                self.zmq_handler.process_queue(wrapper)

                # Check for slide changes
                slide_index = self.slide_tracker.check_slide_change()
                slide_changed = slide_index is not None

                if slide_changed or ServerAction.SEND_CURRENT_VIEWS in requested_actions:
                    target_index = slide_index or self.ppt_controller.get_current_slide_index()
                    images = self.ppt_controller.extract_metadata_images(target_index)
                    parts = self.build_view_message("CurrentViews", images, target_index)
                    self.zmq_handler.send_multipart(parts)

                if ServerAction.SEND_ALL_VIEWS in requested_actions:
                    images = self.ppt_controller.extract_all_metadata_images()
                    parts = self.build_view_message("AllViews", images)
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

    def shutdown(self):
        """Shuts down the ppt_server."""
        self.running = False
        self.zmq_handler.close()
        print("✅ PowerPoint Server shut down successfully.")
        sys.exit(0)
