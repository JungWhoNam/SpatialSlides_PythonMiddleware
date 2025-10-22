import argparse
import signal
import sys
import logging
from ppt_server.ppt_server import PowerPointServer


# Configure logging globally
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)


def main():
    """Starts the PowerPoint ppt_server with CLI-based settings."""

    # Set up argument parser
    parser = argparse.ArgumentParser(description="Start the PowerPoint Server.")

    # Add arguments for configuration
    parser.add_argument("--pub_address", default="tcp://*:5557", help="ZMQ PUB address (default: tcp://*:5557)")
    parser.add_argument("--rep_address", default="tcp://*:5558", help="ZMQ REP address (default: tcp://*:5558)")
    parser.add_argument("--interval", type=float, default=1.0,
                        help="PowerPoint polling interval in seconds (default: 1.0)")

    # Parse CLI arguments
    args = parser.parse_args()

    # Log the starting settings
    logging.info("Starting PowerPoint Server with settings:")
    logging.info(f"   PUB Address: {args.pub_address}")
    logging.info(f"   REP Address: {args.rep_address}")
    logging.info(f"   Polling Interval: {args.interval}s")

    server = PowerPointServer(interval=args.interval)

    # Handle termination signals
    def shutdown_handler(signum, frame):
        """Ensures clean shutdown when stopping the ppt_server."""
        logging.info("Received termination signal. Shutting down...")
        server.shutdown()
        sys.exit(0)

    # Register signal handlers for SIGINT (Ctrl+C) and SIGTERM
    signal.signal(signal.SIGINT, shutdown_handler)
    signal.signal(signal.SIGTERM, shutdown_handler)

    try:
        server.start()
    except KeyboardInterrupt:
        shutdown_handler(None, None)


if __name__ == "__main__":
    main()
