# SAP GUI Debug Tool
# ==================
# Diagnostic wrapper for SAP GUI scripts. When a script fails, captures:
#   - A screenshot of the SAP window at the point of failure
#   - An annotated screenshot with numbered element markers
#   - A text file listing every visible element (ID, type, text, position)
#
# The text file and annotated screenshot use matching numbers so you can
# look at the screenshot to see what's on screen, then find the number
# in the text file to get the exact ID for your script.
#
# Usage:
#   from process_automation import sap_debug
#   sap_debug(r"C:\path\to\template.py", output_dir=r"C:\Temp")
#
# If sap_script succeeds: nothing is saved, runs normally.
# If sap_script fails: saves screenshot + annotated screenshot + element list.

import os
import importlib.util
import traceback
from datetime import datetime

try:
    from PIL import ImageGrab, Image, ImageDraw, ImageFont
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False


# -- Element tree walker -------------------------------------------------------

# SAP GUI element types mapped to friendly names
_TYPE_NAMES = {
    0:  "Unknown",
    10: "Window",
    20: "Dialog",
    21: "ModalDialog",
    30: "Label",
    31: "TextField",
    32: "PasswordField",
    33: "Combobox",
    34: "Checkbox",
    35: "RadioButton",
    36: "Button",
    37: "Tab",
    38: "TabStrip",
    40: "Container",
    41: "SimpleContainer",
    42: "ScrollContainer",
    43: "TableControl",
    44: "TableColumn",
    45: "TableRow",
    46: "TableCell",
    50: "Tree",
    51: "Shell",
    60: "Toolbar",
    61: "Menubar",
    62: "Menu",
    70: "StatusBar",
    71: "StatusPane",
    80: "CustomControl",
    100: "GuiSplit",
    101: "GuiSplitterShell",
    109: "Picture",
    110: "GuiTextedit",
    111: "GuiOfficeIntegration",
    120: "GuiHTMLViewer",
}


def _get_type_name(element):
    """Get a friendly type name for a SAP GUI element."""
    try:
        type_val = element.Type
        name = _TYPE_NAMES.get(type_val, f"Type({type_val})")
    except Exception:
        name = "Unknown"

    # Also try the .TypeAsNumber fallback
    try:
        type_name = element.TypeAsNumber
        if type_name:
            return _TYPE_NAMES.get(type_name, name)
    except Exception:
        pass

    return name


def _get_element_info(element):
    """Extract info from a single SAP GUI element."""
    info = {"id": "", "type": "", "text": "", "left": 0, "top": 0, "width": 0, "height": 0}

    try:
        info["id"] = element.Id
    except Exception:
        pass
    try:
        info["type"] = _get_type_name(element)
    except Exception:
        pass
    try:
        info["text"] = str(element.Text).strip()[:80] if element.Text else ""
    except Exception:
        pass
    try:
        info["left"] = element.ScreenLeft
        info["top"] = element.ScreenTop
        info["width"] = element.Width
        info["height"] = element.Height
    except Exception:
        pass

    return info


def _walk_elements(element, elements=None):
    """Recursively walk the SAP GUI element tree and collect all elements."""
    if elements is None:
        elements = []

    info = _get_element_info(element)
    if info["id"]:
        elements.append(info)

    try:
        children = element.Children
        for i in range(children.Count):
            try:
                _walk_elements(children(i), elements)
            except Exception:
                continue
    except Exception:
        pass

    return elements


def _get_container_label(element_id):
    """Extract a container group label from a full element ID."""
    # e.g. "wnd[0]/tbar[0]/btn[3]" -> "Toolbar [wnd[0]/tbar[0]]"
    # e.g. "wnd[0]/usr/ctxtFOO" -> "User Area [wnd[0]/usr]"
    # e.g. "wnd[0]/mbar/menu[0]" -> "Menu Bar [wnd[0]/mbar]"
    parts = element_id.split("/")
    if len(parts) < 3:
        return "Window"

    # Take everything up to the second-to-last part as the container
    container_path = "/".join(parts[:-1])
    last_container = parts[-2] if len(parts) >= 2 else ""

    if "tbar" in last_container:
        return f"Toolbar [{container_path}]"
    elif "usr" in container_path:
        return f"User Area [{container_path}]"
    elif "mbar" in last_container or "menu" in last_container:
        return f"Menu Bar [{container_path}]"
    elif "sbar" in last_container:
        return f"Status Bar [{container_path}]"
    elif "titl" in last_container:
        return f"Title Bar [{container_path}]"
    else:
        return f"[{container_path}]"


# -- Screenshot ----------------------------------------------------------------

def _take_screenshot(session):
    """Take a screenshot of the SAP window. Returns a PIL Image or None."""
    if not _HAS_PIL:
        return None

    try:
        wnd = session.findById("wnd[0]")
        left = wnd.ScreenLeft
        top = wnd.ScreenTop
        width = wnd.Width
        height = wnd.Height
        img = ImageGrab.grab(bbox=(left, top, left + width, top + height))
        return img
    except Exception:
        # Fallback: grab full screen
        try:
            return ImageGrab.grab()
        except Exception:
            return None


def _annotate_screenshot(img, elements, window_left, window_top):
    """Draw numbered markers on the screenshot matching element positions."""
    annotated = img.copy()
    draw = ImageDraw.Draw(annotated)

    try:
        font = ImageFont.truetype("arial.ttf", 12)
    except Exception:
        font = ImageFont.load_default()

    for i, elem in enumerate(elements, 1):
        # Convert screen coordinates to image coordinates
        x = elem["left"] - window_left
        y = elem["top"] - window_top

        if x < 0 or y < 0 or x > img.width or y > img.height:
            continue

        # Draw a small red circle with the number
        radius = 8
        draw.ellipse(
            [x - radius, y - radius, x + radius, y + radius],
            fill="red", outline="white", width=1,
        )
        draw.text((x - 4, y - 6), str(i), fill="white", font=font)

    return annotated


# -- Text dump -----------------------------------------------------------------

def _build_element_dump(elements, failed_id=None, window_title=""):
    """Build a formatted text dump of all elements, grouped by container."""
    lines = []
    lines.append(f"SCREEN: {window_title}")
    if failed_id:
        lines.append(f"FAILED AT: {failed_id}")
    lines.append(f"TOTAL ELEMENTS: {len(elements)}")
    lines.append("")

    # Group elements by container
    groups = {}
    for i, elem in enumerate(elements, 1):
        container = _get_container_label(elem["id"])
        if container not in groups:
            groups[container] = []
        groups[container].append((i, elem))

    for container, items in groups.items():
        # Sort by top, then left
        items.sort(key=lambda x: (x[1]["top"], x[1]["left"]))

        lines.append(f"── {container} " + "─" * max(0, 60 - len(container)))
        for num, elem in items:
            short_id = elem["id"].split("/")[-1] if "/" in elem["id"] else elem["id"]
            text_preview = f'"{elem["text"]}"' if elem["text"] else ""
            pos = f'({elem["left"]}, {elem["top"]})'
            lines.append(f' {num:>3}. {short_id:<25} {elem["type"]:<15} {text_preview:<40} {pos}')
        lines.append("")

    return "\n".join(lines)


# -- Main debug function -------------------------------------------------------

def _load_sap_script(script_path):
    """Load sap_script function from a template file."""
    spec = importlib.util.spec_from_file_location("template", script_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    if not hasattr(mod, "sap_script"):
        raise AttributeError(f"No 'sap_script' function found in {script_path}")
    return mod.sap_script


def sap_debug(script_path, output_dir=None):
    """
    Run a SAP template with diagnostic capture on failure.

    Opens a SAP session, loads sap_script from the template file, runs it,
    and closes the session. If the script fails, captures:
      - A screenshot of the SAP window
      - An annotated screenshot with numbered element markers
      - A text file listing every visible element

    Parameters
    ----------
    script_path : path to a SAP template .py file containing a sap_script(session) function
    output_dir  : where to save debug files (default: user's Desktop)
    """
    from .sap_connection import SAPManager

    if output_dir is None:
        output_dir = os.path.join(os.environ.get("USERPROFILE", "."), "Desktop")

    sap_script = _load_sap_script(script_path)

    sap = SAPManager()
    session = sap.get_session()

    try:
        sap_script(session)
        print("[SAP DEBUG] Script completed successfully.")
        return True
    except Exception as e:
        print(f"\n[SAP DEBUG] Script failed: {e}")

        # Extract the failed findById ID from the traceback if possible
        failed_id = None
        tb_text = traceback.format_exc()
        if "findById" in tb_text:
            for line in tb_text.splitlines():
                if "findById" in line:
                    start = line.find('findById("') + len('findById("')
                    end = line.find('"', start)
                    if start > 0 and end > start:
                        failed_id = line[start:end]
                    break

        # Get window title
        window_title = ""
        try:
            window_title = session.findById("wnd[0]").text
        except Exception:
            pass

        # Walk the element tree
        print("[SAP DEBUG] Walking element tree...")
        try:
            wnd = session.findById("wnd[0]")
            elements = _walk_elements(wnd)
        except Exception as walk_err:
            print(f"[SAP DEBUG] Could not walk element tree: {walk_err}")
            elements = []

        # Filter to actionable elements (skip pure containers)
        actionable = [e for e in elements if e["type"] not in ("Window", "Container", "SimpleContainer", "ScrollContainer", "Unknown")]

        # Take screenshot
        print("[SAP DEBUG] Capturing screenshot...")
        screenshot = _take_screenshot(session)

        os.makedirs(output_dir, exist_ok=True)

        # Save annotated screenshot
        if screenshot:
            try:
                window_left = session.findById("wnd[0]").ScreenLeft
                window_top = session.findById("wnd[0]").ScreenTop
                annotated = _annotate_screenshot(screenshot, actionable, window_left, window_top)
                img_path = os.path.join(output_dir, "sap_debug.png")
                annotated.save(img_path)
                print(f"[SAP DEBUG] Screenshot: {img_path}")
            except Exception as ann_err:
                print(f"[SAP DEBUG] Could not annotate screenshot: {ann_err}")
        else:
            print("[SAP DEBUG] Screenshot not available (install Pillow: pip install Pillow)")

        # Save element dump
        dump = _build_element_dump(actionable, failed_id, window_title)
        txt_path = os.path.join(output_dir, "sap_debug.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(dump)
        print(f"[SAP DEBUG] Elements:   {txt_path}")

        # Also print a summary to terminal
        print(f"\n{dump}")

        return False
    finally:
        try:
            sap.close_connection(session)
        except Exception:
            pass
