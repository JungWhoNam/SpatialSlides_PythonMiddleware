import tempfile
import os


def handle_message(message_parts, ppt_controller):
    """
    Processes multipart messages from Unity.
    [ JSON metadata (bytes), Image data (bytes) ]
    """

    if len(message_parts) < 2:
        print("⚠️ Received incomplete message.")
        return

    metadata_json = message_parts[0].decode("utf-8")
    image_bytes = message_parts[1]

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
        temp_file.write(image_bytes)
        temp_file_path = temp_file.name

    ppt_controller.insert_metadata_image_offscreen(temp_file_path, metadata_json, apply_style=True)

    try:
        os.remove(temp_file_path)
    except Exception as e:
        print(f"⚠️ Error deleting temp file: {e}")

    print(f"📄 Inserted image with metadata: {metadata_json}")
