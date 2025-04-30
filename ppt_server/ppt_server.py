import sys
import time
from .zmq_handler import ZMQHandler
from ppt.ppt_controller import PowerPointController
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
        if not self.ppt_controller.connect_to_powerpoint():
            print("❌ Could not connect to PowerPoint. Exiting.")
            return

        print("✅ PowerPoint Server is running. Processing incoming messages...")

        try:
            while self.running:
                start_time = time.time()

                # Process incoming messages from Unity
                self.zmq_handler.process_queue(handle_message, self.ppt_controller)

                # Check for slide changes
                change = self.slide_tracker.check_slide_change()
                if change:
                    slide_index, notes = change
                    self.zmq_handler.send_message("SlideChanged", {"slide": slide_index, "notes": notes})

                # Ensure at least some small sleep to prevent high CPU usage
                elapsed_time = time.time() - start_time
                time.sleep(max(0, self.interval - elapsed_time))

        except KeyboardInterrupt:
            self.shutdown()

    def shutdown(self):
        """Shuts down the ppt_server."""
        self.running = False
        self.zmq_handler.close()
        print("✅ PowerPoint Server shut down successfully.")
        sys.exit(0)
