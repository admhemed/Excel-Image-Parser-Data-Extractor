import os
from typing import List, Dict, Optional

import win32com.client as win32

# ==========================
# Paths / Settings
# ==========================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.join(SCRIPT_DIR, "TurnKeyFiles2")
IMAGES_DIR = os.path.join(SCRIPT_DIR, "images")

# msoShapeType constants (we mainly care about pictures)
MSO_PICTURE = 13  # from Office constants


# ==========================
# Helpers
# ==========================

def ensure_dir(path: str) -> None:
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)


def is_excel_file(path: str) -> bool:
    name = os.path.basename(path)
    if name.startswith("~$"):
        return False
    return name.lower().endswith(".xlsx")


def build_row_boundaries_excel(ws) -> List[Dict]:
    """
    Ask Excel for each row's Top & Height (in points).

    Returns list of dicts:
    [
      {"row": 1, "y_top": 0.0, "y_bottom": ...},
      ...
    ]
    """
    boundaries: List[Dict] = []

    used_range = ws.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1

    for row_idx in range(first_row, last_row + 1):
        row_obj = ws.Rows(row_idx)
        top = float(row_obj.Top)        # points
        height = float(row_obj.Height)  # points
        y_top = top
        y_bottom = top + height
        boundaries.append({
            "row": row_idx,
            "y_top": y_top,
            "y_bottom": y_bottom,
        })

    return boundaries


def find_row_for_y(center_y: float, row_boundaries: List[Dict]) -> Optional[int]:
    """
    Given a Y (points), return the row index that contains this Y.
    Returns None if no row range matches.
    """
    for rb in row_boundaries:
        if rb["y_top"] <= center_y < rb["y_bottom"]:
            return rb["row"]
    return None


def export_shape_to_image(shp, global_index: int) -> str:
    """
    Export the shape as a PNG image to IMAGES_DIR with a name based on the global index.
    Returns the full path to the saved image.
    """
    ensure_dir(IMAGES_DIR)
    filename = f"img_{global_index:04d}.png"
    full_path = os.path.join(IMAGES_DIR, filename)

    try:
        # Direct Export from shape (Excel COM)
        shp.Export(Filename=full_path, FilterName="PNG")
    except Exception as e:
        print(f"    [WARN] Could not export shape '{shp.Name}' as image: {e}")
        # create empty file so that index stays consistent
        with open(full_path, "wb") as f:
            f.write(b"")

    return full_path


# ==========================
# Process one workbook
# ==========================

def process_workbook(excel_app, path: str, global_index_start: int) -> int:
    """
    Process a single workbook.
    Returns the next global index after processing all shapes in this workbook.
    """
    rel_path = os.path.relpath(path, ROOT_DIR)
    print(f"\n[INFO] =============================================")
    print(f"[INFO] Processing workbook: {rel_path}")

    wb = excel_app.Workbooks.Open(path, ReadOnly=True)
    current_index = global_index_start

    try:
        for ws in wb.Worksheets:
            sheet_name = ws.Name
            print(f"\n[INFO]  Sheet: {sheet_name}")

            shapes = ws.Shapes
            try:
                shapes_count = shapes.Count
            except Exception as e:
                print(f"[ERROR]   Cannot read Shapes.Count: {e}")
                continue

            print(f"[INFO]   Shapes.Count = {shapes_count}")

            if shapes_count == 0:
                print("[INFO]   No shapes in this sheet.")
                continue

            # Build row boundaries once for this sheet (for center-based rows)
            row_boundaries = build_row_boundaries_excel(ws)

            for i in range(1, shapes_count + 1):
                try:
                    shp = shapes.Item(i)
                except Exception as e:
                    print(f"  [SHAPE #{i}] ERROR: cannot get shape: {e}")
                    continue

                try:
                    shp_type = shp.Type
                except Exception:
                    shp_type = None

                # We care about pictures; you can loosen this if needed
                if shp_type != MSO_PICTURE:
                    continue

                current_index += 1  # assign a unique global index

                name = shp.Name
                left = float(shp.Left)
                top = float(shp.Top)
                width = float(shp.Width)
                height = float(shp.Height)
                center_y = top + height / 2.0

                # Get row_start / row_end from Excel via TopLeftCell / BottomRightCell
                row_start = row_end = anchor_row = None
                try:
                    top_left_cell = shp.TopLeftCell
                    bottom_right_cell = shp.BottomRightCell
                    row_start = int(top_left_cell.Row)
                    row_end = int(bottom_right_cell.Row)
                except Exception as e:
                    print(f"  [IMAGE #{current_index}] WARNING: cannot get TopLeftCell/BottomRightCell: {e}")

                # Decide final anchor_row
                if row_start is not None and row_end is not None:
                    if row_start == row_end:
                        # Image is considered "on a single cell"
                        anchor_row = row_start
                    else:
                        # Floating / spanning multiple rows:
                        # use center_y + row_boundaries to find row
                        row_by_center = find_row_for_y(center_y, row_boundaries)
                        anchor_row = row_by_center

                # Export shape to image file with this global index
                image_path = export_shape_to_image(shp, current_index)
                rel_image_path = os.path.relpath(image_path, SCRIPT_DIR)

                print(f"  [IMAGE #{current_index}]")
                print(f"    file        = {rel_path}")
                print(f"    sheet       = {sheet_name}")
                print(f"    shape_name  = {name}")
                print(f"    shape_type  = {shp_type}")
                print(f"    image_path  = {rel_image_path}")
                print(f"    left        = {left}")
                print(f"    top         = {top}")
                print(f"    width       = {width}")
                print(f"    height      = {height}")
                print(f"    center_y    = {center_y}")

                if row_start is None or row_end is None:
                    print(f"    row_start   = UNKNOWN")
                    print(f"    row_end     = UNKNOWN")
                else:
                    print(f"    row_start   = {row_start}")
                    print(f"    row_end     = {row_end}")

                if anchor_row is None:
                    print(f"    anchor_row  = UNKNOWN (center-based search did not find a row)")
                else:
                    print(f"    anchor_row  = {anchor_row}")

                print("")

    finally:
        wb.Close(SaveChanges=False)

    return current_index


# ==========================
# Walk all workbooks
# ==========================

def walk_and_process_root(root_dir: str) -> None:
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.ScreenUpdating = False
    excel_app.DisplayAlerts = False

    global_index = 0

    try:
        for dirpath, dirnames, filenames in os.walk(root_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                if not is_excel_file(full_path):
                    continue

                try:
                    global_index = process_workbook(excel_app, full_path, global_index)
                except Exception as e:
                    rel = os.path.relpath(full_path, root_dir)
                    print(f"[ERROR] Failed to process '{rel}': {e}")
    finally:
        excel_app.Quit()


if __name__ == "__main__":
    print(f"[INFO] SCRIPT_DIR = {SCRIPT_DIR}")
    print(f"[INFO] ROOT_DIR   = {ROOT_DIR}")
    print(f"[INFO] IMAGES_DIR = {IMAGES_DIR}")
    walk_and_process_root(ROOT_DIR)
