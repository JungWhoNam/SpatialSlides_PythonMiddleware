import os
from ppt_tools import PowerPointController


def format_bytes(size: int) -> str:
    """Convert byte size into human-readable format."""
    for unit in ['bytes', 'KB', 'MB', 'GB']:
        if size < 1024.0:
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"


def main():
    controller = PowerPointController()
    if not controller.connect():
        print("❌ Could not connect to PowerPoint.")
        return

    # Insert image just outside the slide with metadata
    # image_path = os.path.join(os.getcwd(), "run_demo.png")
    # metadata = '{"label": "test-image"}'
    # controller.insert_metadata_image_offscreen(image_path, metadata, apply_style=True)

    slide_index = controller.get_current_slide_index()
    if slide_index is None:
        print("⚠️ Could not determine current slide.")
        return

    print("\n📦 Extracted images with JSON metadata:")
    for i, (img_bytes, alt) in enumerate(controller.extract_metadata_images(slide_index)):
        print(f"  Image {i + 1}: {format_bytes(len(img_bytes))}, Alt Text: {alt}")


if __name__ == "__main__":
    main()
