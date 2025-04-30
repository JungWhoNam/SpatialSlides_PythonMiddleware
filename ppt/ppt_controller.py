import win32com.client
from typing import Optional, List, Tuple


class PowerPointController:
    def __init__(self) -> None:
        self.app: Optional[win32com.client.CDispatch] = None

    def connect_to_powerpoint(self) -> bool:
        """Connects to an active PowerPoint instance."""
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            if self.app.Presentations.Count == 0:
                print("❌ No open PowerPoint presentations found.")
                return False
            return True
        except Exception as e:
            print(f"❌ Error connecting to PowerPoint: {e}")
            return False

    def get_current_slide_index(self) -> Optional[int]:
        """Returns current slide index (1-based)."""
        if not self.app:
            return None
        try:
            return self.app.ActiveWindow.View.Slide.SlideIndex
        except Exception:
            return None

    def _get_presentation(self) -> Optional[win32com.client.CDispatch]:
        """Returns active presentation object."""
        if self.app:
            try:
                return self.app.ActivePresentation
            except Exception:
                return None
        return None

    def _get_slide(self, index: Optional[int] = None) -> Optional[win32com.client.CDispatch]:
        """Returns slide object for given index or current slide."""
        pres = self._get_presentation()
        if not pres:
            return None
        try:
            index = index or self.get_current_slide_index()
            if not index or index < 1 or index > pres.Slides.Count:
                return None
            return pres.Slides(index)
        except Exception:
            return None

    def _get_slide_dimensions(self) -> Tuple[float, float]:
        """Returns width and height of slide."""
        pres = self._get_presentation()
        if not pres:
            return (0, 0)
        return pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight

    def _add_image(
            self,
            slide: win32com.client.CDispatch,
            image_path: str,
            left: float,
            top: float,
            width: float,
            height: float,
            alt_text: str = "",
            apply_style: bool = True
    ) -> Optional[win32com.client.CDispatch]:
        """Inserts an image on a slide with optional border styling and alt text."""
        try:
            image = slide.Shapes.AddPicture(
                FileName=image_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=left,
                Top=top,
                Width=width,
                Height=height
            )
            if alt_text:
                image.AlternativeText = alt_text
            if apply_style:
                image.Line.Visible = True
                image.Line.Weight = 3
                image.Line.ForeColor.RGB = 21 + (96 << 8) + (130 << 16)  # RGB(130, 96, 21)
            return image
        except Exception as e:
            print(f"❌ Failed to insert image: {e}")
            return None

    def get_image_alt_text_from_slide(self, slide_index: int) -> Optional[List[str]]:
        """Returns list of alt texts from all image shapes on given slide."""
        slide = self._get_slide(slide_index)
        if not slide:
            print("⚠️ Could not access specified slide.")
            return None

        alt_texts = []
        for shape in slide.Shapes:
            if shape.Type == 13:  # msoPicture
                try:
                    alt = shape.AlternativeText.strip()
                    if alt:
                        alt_texts.append(alt)
                except Exception:
                    alt_texts.append("⚠️ Unable to retrieve Alt Text")
        return alt_texts or None

    def add_empty_slide_with_text(self, text: str) -> None:
        """Creates a blank slide and places a text box in the center."""
        pres = self._get_presentation()
        if not pres:
            print("⚠️ PowerPoint is not connected.")
            return
        try:
            slide = pres.Slides.Add(pres.Slides.Count + 1, 12)  # 12 = ppLayoutBlank
            textbox = slide.Shapes.AddTextbox(1, 100, 100, 500, 50)
            textbox.TextFrame.TextRange.Text = text
            print(f"📄 Added empty slide with text: {text}")
        except Exception as e:
            print(f"❌ Error adding slide: {e}")

    def add_empty_slide_with_temp_image(self, image_path: str, metadata_json: str = "",
                                        apply_style: bool = True) -> None:
        """Adds a blank slide and places an image at the center."""
        pres = self._get_presentation()
        if not pres:
            print("⚠️ PowerPoint is not connected.")
            return
        try:
            slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
            width, height = self._get_slide_dimensions()
            img_w, img_h = width * 0.6, height * 0.6
            left = (width - img_w) / 2
            top = (height - img_h) / 2
            self._add_image(slide, image_path, left, top, img_w, img_h, metadata_json, apply_style)
            print(f"🖼️ Slide with image and metadata added.")
        except Exception as e:
            print(f"❌ Error adding slide with image: {e}")

    def add_corner_image_to_current_slide(self, image_path: str, metadata_json: str = "",
                                          apply_style: bool = True) -> None:
        """Places a small image just outside top-right of slide. Moves down if space is taken."""
        slide = self._get_slide()
        if not slide:
            print("❌ Failed to retrieve current slide.")
            return

        width, height = self._get_slide_dimensions()
        img_w, img_h = width * 0.1, height * 0.1

        # Offset to avoid overlapping previously added images
        margin = 10
        count = sum(
            1 for s in slide.Shapes
            if s.Type == 13 and s.Left > width
        )
        left = width + margin
        top = -margin + (count * (img_h + margin))

        self._add_image(slide, image_path, left, top, img_w, img_h, metadata_json, apply_style)
        print("🖼️ Corner image added outside slide.")
