import win32com.client
from typing import Optional, Tuple, List


def get_active_presentation(app: win32com.client.CDispatch) -> Optional[win32com.client.CDispatch]:
    try:
        return app.ActivePresentation
    except Exception:
        return None


def is_presenter_mode(app: win32com.client.CDispatch) -> Optional[bool]:
    try:
        return app.SlideShowWindows.Count > 0
    except Exception as e:
        print(f"⚠️ Error checking presenter mode: {e}")
        return None


def get_current_slide_index(app: win32com.client.CDispatch) -> Optional[int]:
    try:
        if is_presenter_mode(app):
            return app.SlideShowWindows(1).View.Slide.SlideIndex
        else:
            return app.ActiveWindow.View.Slide.SlideIndex
    except Exception:
        return None


def get_current_click_index(app: win32com.client.CDispatch) -> Optional[int]:
    """
    Gets the current animation step (click index) on the active slide.
    Returns the click index (an integer, starting at 0) if in presenter mode.
    Returns None if in edit mode or if an error occurs.
    """
    try:
        if is_presenter_mode(app):
            return app.SlideShowWindows(1).View.GetClickIndex()
        else:
            return None
    except Exception:
        return None


def get_slide_by_index(app: win32com.client.CDispatch, index: Optional[int] = None) -> Optional[
    win32com.client.CDispatch]:
    pres = get_active_presentation(app)
    if not pres:
        return None
    try:
        if index is None:
            index = get_current_slide_index(app)
        if index is None or index < 1 or index > pres.Slides.Count:
            return None
        return pres.Slides(index)
    except Exception:
        return None


def get_presentation_dimensions(app: win32com.client.CDispatch) -> Tuple[float, float]:
    pres = get_active_presentation(app)
    if not pres:
        return (0, 0)
    return pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight


def collect_image_alt_texts(app: win32com.client.CDispatch, slide_index: int) -> Optional[List[str]]:
    slide = get_slide_by_index(app, slide_index)
    if not slide:
        print("⚠️ Could not access specified slide.")
        return None

    alt_texts = []
    for shape in slide.Shapes:
        if shape.Type == 13:
            try:
                text = shape.AlternativeText.strip()
                if text:
                    alt_texts.append(text)
            except Exception:
                alt_texts.append("⚠️ Unable to retrieve Alt Text")
    return alt_texts or None


def create_slide_with_textbox(app: win32com.client.CDispatch, text: str) -> None:
    pres = get_active_presentation(app)
    if not pres:
        print("⚠️ PowerPoint is not connected.")
        return

    try:
        slide_count = pres.Slides.Count
        new_slide = pres.Slides.Add(slide_count + 1, 12)
        text_box = new_slide.Shapes.AddTextbox(1, 100, 100, 500, 50)
        text_box.TextFrame.TextRange.Text = text
        print(f"📄 Added empty slide at position {slide_count + 1} with text: {text}")
    except Exception as e:
        print(f"❌ Error adding slide: {e}")
