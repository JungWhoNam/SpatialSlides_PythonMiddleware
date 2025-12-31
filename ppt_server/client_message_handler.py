import tempfile
import os
import logging

from .server_action import ServerAction


def handle_message(message_parts, ppt_controller) -> ServerAction:
    """
    Processes multipart messages from Unity.

    message_parts:
      - ["GetCurrentViews"] → SEND_CURRENT_VIEWS
      - ["CreateView", <metadata>, <image>] → Inserts image in PPT and returns NO_ACTION

    Returns:
        ServerAction enum indicating server behavior.
    """

    if not message_parts or len(message_parts) == 0:
        logging.warning("Empty message received.")
        return ServerAction.NO_ACTION

    try:
        command = message_parts[0].decode("utf-8")
    except UnicodeDecodeError:
        logging.warning("Failed to decode message header.")
        return ServerAction.NO_ACTION

    if command == "GetCurrentViews":
        return ServerAction.SEND_CURRENT_VIEWS

    elif command == "GetAllViews":
        return ServerAction.SEND_ALL_VIEWS

    elif command == "GetCurrentMode":
        return ServerAction.SEND_CURRENT_MODE

    elif command == "CreateView":
        temp_path = None
        if len(message_parts) < 3:
            logging.warning("CreateView requires 2 parts: metadata and image.")
            return ServerAction.NO_ACTION

        try:
            metadata_json: str = message_parts[1].decode("utf-8")
            image_bytes: bytes = message_parts[2]

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                temp_file.write(image_bytes)
                temp_path: str = temp_file.name

            ppt_controller.insert_metadata_image_offscreen(temp_path, metadata_json, apply_style=True)

            return ServerAction.SEND_CURRENT_VIEWS
        except Exception as e:
            logging.error(f"Error handling CreateView: {e}")
        finally:
            try:
                if temp_path and os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception as e:
                logging.warning(f"Error deleting temp file: {e}")

    else:
        logging.warning(f"Unknown command: {command}")

    return ServerAction.NO_ACTION