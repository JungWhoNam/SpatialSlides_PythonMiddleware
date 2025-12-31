import tempfile
import os
import json
from typing import Optional, List, Tuple
import win32com.client
import logging

from ppt_tools.utils_slide import get_presentation_dimensions, get_active_presentation


def insert_image(
        slide: win32com.client.CDispatch,
        image_path: str,
        left: float,
        top: float,
        width: float,
        height: float,
        alt_text: str = "",
        apply_style: bool = True
) -> Optional[win32com.client.CDispatch]:
    try:
        # Ensure image_path is absolute for COM reliability
        abs_image_path = os.path.abspath(image_path)

        image_shape = slide.Shapes.AddPicture(
            FileName=abs_image_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=width,
            Height=height
        )
        if alt_text:
            image_shape.AlternativeText = alt_text

        if apply_style:
            image_shape.Line.Visible = True
            image_shape.Line.Weight = 3
            image_shape.Line.ForeColor.RGB = 21 + (96 << 8) + (130 << 16)

        return image_shape
    except Exception as e:
        logging.error(f"Failed to insert image from path '{image_path}': {e}")
        return None


def extract_images_with_json_metadata(slide: win32com.client.CDispatch) -> List[Tuple[bytes, Optional[str]]]:
    result: List[Tuple[bytes, Optional[str]]] = []

    for shape in slide.Shapes:
        if shape.Type == 13:
            alt_text = shape.AlternativeText.strip() if shape.AlternativeText else None

            # 1. Validate Alt Text is non-empty and parsable JSON
            try:
                if not alt_text:
                    continue
                json.loads(alt_text)
            except json.JSONDecodeError as e:
                logging.warning(f"Skipping image with non-JSON Alt Text: {alt_text}. Error: {e}")
                continue
            except Exception as e:
                logging.warning(f"Skipping shape due to unexpected error during AltText access: {e}")
                continue

            # 2. Export image to temp file, read bytes, and delete
            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp_path = tmp.name
                shape.Export(tmp_path, 2)
                with open(tmp_path, "rb") as f:
                    image_bytes = f.read()
                os.remove(tmp_path)
                result.append((image_bytes, alt_text))
            except Exception as e:
                logging.error(f"Failed to export or read image data from slide: {e}")
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except Exception as e_del:
                        logging.warning(f"Failed to delete temporary file {tmp_path}: {e_del}")

    return result


def extract_json_metadata_only(slide: win32com.client.CDispatch) -> List[str]:
    """
    Extracts only the JSON metadata (from Alt Text) for valid images,
    sorted by their vertical position on the slide.
    """
    metadata_with_positions = []
    for shape in slide.Shapes:
        if shape.Type == 13:
            alt_text = shape.AlternativeText.strip() if shape.AlternativeText else None
            try:
                if not alt_text:
                    continue
                json.loads(alt_text)
                metadata_with_positions.append({
                    "top": shape.Top,
                    "alt": alt_text
                })
            except json.JSONDecodeError as e:
                logging.warning(f"Skipping shape due to invalid JSON metadata: {e}")
                continue
            except Exception as e:
                logging.warning(f"Skipping shape due to COM error during metadata access: {e}")
                continue

    sorted_metadata = sorted(metadata_with_positions, key=lambda meta: meta['top'])

    return [meta['alt'] for meta in sorted_metadata]


def create_slide_with_centered_image(
        app: win32com.client.CDispatch,
        image_path: str,
        metadata_json: str = "",
        apply_style: bool = True
) -> None:
    pres = get_active_presentation(app)
    if not pres:
        logging.warning("PowerPoint is not connected.")
        return

    try:
        slide_count = pres.Slides.Count
        new_slide = pres.Slides.Add(slide_count + 1, 12)

        width, height = get_presentation_dimensions(app)
        image_width = width * 0.6
        image_height = height * 0.6
        left = (width - image_width) / 2
        top = (height - image_height) / 2

        insert_image(new_slide, image_path, left, top, image_width, image_height, metadata_json, apply_style)
        logging.info(f"Slide {slide_count + 1} created with a centered image and metadata.")
    except Exception as e:
        logging.error(f"Error adding slide with image: {e}")