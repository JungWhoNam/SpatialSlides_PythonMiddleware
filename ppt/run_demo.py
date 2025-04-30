from ppt_controller import PowerPointController
import os


def main() -> None:
    controller = PowerPointController()

    if not controller.connect_to_powerpoint():
        print("🔌 Could not connect to PowerPoint.")
        return

    slide_index = controller.get_current_slide_index()
    if slide_index is None:
        print("❌ Could not get current slide index.")
        return

    print(f"📊 Current slide index: {slide_index}")

    image_path = os.path.join(os.getcwd(), "run_demo.png")
    metadata = '{"source": "test", "desc": "corner test image"}'
    controller.add_corner_image_to_current_slide(image_path, metadata, apply_style=True)

    alt_texts = controller.get_image_alt_text_from_slide(slide_index)
    if alt_texts:
        print("📝 Alt Texts on this slide:")
        for i, alt in enumerate(alt_texts, 1):
            print(f"  {i}. {alt}")
    else:
        print("ℹ️ No Alt Texts found on this slide.")


if __name__ == "__main__":
    main()
