import zmq
import queue
import threading
import json


class ZMQHandler:
    """Handles ZMQ communication between PowerPoint and external clients (Unity)."""

    def __init__(self, pub_address="tcp://*:5557", pull_address="tcp://*:5558"):
        """
        Initializes ZMQ communication:
        - PUB socket for broadcasting messages (PUB-SUB)
        - PULL socket for receiving client requests (PUSH-PULL)
        """
        self.context = zmq.Context()

        # Publisher socket for broadcasting events (PUB-SUB)
        self.publisher = self.context.socket(zmq.PUB)
        self.publisher.bind(pub_address)

        # PULL socket for handling incoming messages (PUSH-PULL)
        self.pull_socket = self.context.socket(zmq.PULL)
        self.pull_socket.bind(pull_address)

        # Message Queue (Thread-Safe)
        self.message_queue = queue.Queue()

        # Thread control flags
        self.running = True
        self.request_thread = threading.Thread(target=self.listen_for_requests, daemon=True)
        self.request_thread.start()

        print(f"✅ ZMQ Server running. PUB: {pub_address}, PULL: {pull_address}")

    def send_message(self, event: str, data: any) -> None:
        """Sends a JSON message to all subscribers."""
        if self.running:
            message = {"event": event, "data": data}
            try:
                self.publisher.send_json(message)
                print(f"📤 Published: {message}")
            except zmq.ZMQError as e:
                print(f"⚠️ Error sending message: {e}")

    def listen_for_requests(self):
        """Listens for incoming multipart messages and adds them to the processing queue."""
        print("🎧 Listening for Unity messages...")

        while self.running:
            try:
                # Receive multipart messages
                message_parts = self.pull_socket.recv_multipart()

                if not message_parts:
                    print("⚠️ Received empty message.")
                    continue

                # Add received data to the queue
                self.message_queue.put(message_parts)

                print(f"📩 Received multipart message: {len(message_parts)} parts")

            except zmq.ZMQError as e:
                if self.running:
                    print(f"⚠️ ZMQ Error receiving request: {e}")

            except Exception as e:
                if self.running:
                    print(f"⚠️ Unexpected error: {e}")

    def process_queue(self, handler, *args):
        """Processes queued messages using an external handler function."""
        while not self.message_queue.empty():
            try:
                message_parts = self.message_queue.get_nowait()
                handler(message_parts, *args)
            except queue.Empty:
                break

    def close(self):
        """Stops ZMQ communication and ensures all threads terminate properly."""
        print("🛑 Closing ZMQHandler...")
        self.running = False

        # Close ZMQ sockets properly
        if self.publisher:
            self.publisher.close(linger=0)
        if self.pull_socket:
            self.pull_socket.close(linger=0)

        self.context.term()

        # Ensure the request thread stops
        if self.request_thread.is_alive():
            self.request_thread.join(timeout=1)

        print("✅ ZMQHandler shut down successfully.")
