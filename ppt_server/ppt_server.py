import sys
import time
import json
from typing import Optional
from .zmq_handler import ZMQHandler
from ppt_tools.powerpoint_controller import PowerPointController
from .slide_tracker import SlideTracker
from .client_message_handler import handle_message


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
                self.zmq_handler.process_queue(handle_message, self.ppt_controller)

                # Check for slide changes
                slide_index = self.slide_tracker.check_slide_change()
                if slide_index:
                    images = self.ppt_controller.extract_metadata_images(slide_index)
                    parts = self.build_slide_changed_message(slide_index, images)
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
