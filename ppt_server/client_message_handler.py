import tempfile


def handle_message(message_parts, ppt_watcher):
    """
    Processes multipart messages from Unity. Supports:
      - Camera metadata (JSON)
      - Image data (PNG bytes)

    Args:
        message_parts (list): List containing multipart message data.
        ppt_watcher (PowerPointWatcher): Instance for managing PowerPoint slides.

    Expected multipart format:
        [ JSON metadata (bytes), Image data (bytes) ]
    """

    # Check if the message contains both JSON metadata and an image
    if len(message_parts) < 2:
        print("⚠️ Received incomplete message.")
        return

    # Decode the first part as JSON metadata
    metadata_json = message_parts[0].decode("utf-8")

    # Convert image bytes to a temporary file
    image_bytes = message_parts[1]
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
        temp_file.write(image_bytes)
        temp_file_path = temp_file.name  # Get temporary file path

    # Add a new slide with the temporary image
    ppt_watcher.add_corner_image_to_current_slide(temp_file_path, metadata_json, True)

    # Delete temporary file after use
    try:
        import os
        os.remove(temp_file_path)
    except Exception as e:
        print(f"⚠️ Error deleting temp file: {e}")

    print(f"📄 Created a new slide with text: {metadata_json} and inserted temp image.")
