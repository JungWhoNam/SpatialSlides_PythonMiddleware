import sys
import time
import json
from typing import Optional
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
                action = ServerAction.NO_ACTION

                def wrapper(msg_parts):
                    nonlocal action
                    result = handle_message(msg_parts, self.ppt_controller)
                    if result != ServerAction.NO_ACTION:
                        action = result

                self.zmq_handler.process_queue(wrapper)

                # Check for slide changes
                slide_index = self.slide_tracker.check_slide_change()
                slide_changed = slide_index is not None

                if slide_changed or action == ServerAction.SEND_CURRENT_VIEWS:
                    target_index = slide_index or self.ppt_controller.get_current_slide_index()
                    images = self.ppt_controller.extract_metadata_images(target_index)
                    parts = self.build_slide_changed_message(target_index, images)
                    self.zmq_handler.send_multipart(parts)

                time.sleep(max(0, self.interval - (time.time() - start_time)))

        except KeyboardInterrupt:
            self.shutdown()

    def build_slide_changed_message(self, slide_index: int, images: list[tuple[bytes, Optional[str]]]) -> list[bytes]:
        """Constructs a multipart message with slide index and associated images."""
        parts = [
            b"SlideChanged",
            json.dumps({"slide": slide_index}).encode("utf-8")
        ]
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
