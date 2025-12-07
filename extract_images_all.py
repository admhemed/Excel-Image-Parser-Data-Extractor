import os
from uuid import uuid4
from collections import Counter
from typing import List, Dict, Any

from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as XLImage

# ==========================
# إعدادات عامة
# ==========================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# الروت المستهدف الذي يحتوي ملفات الإكسل
ROOT_DIR = os.path.join(SCRIPT_DIR, "2025-12-05/Electric")

# فولدر الصور (داخل الروت)
IMAGES_DIR = os.path.join(SCRIPT_DIR, "images")

# اسم ملف الداتا الناتج (داخل الروت أيضاً)
OUTPUT_EXCEL = os.path.join(SCRIPT_DIR, "packages_data.xlsx")

# كلمات الهيدر التي نبحث عنها (تُقارن بحروف صغيرة)
HEADER_KEYWORDS = ("part number", "description", "qty")

# قيم إكسل الخطأ التي لا يجب اعتبارها أسماء باكجات
EXCLUDED_PACKAGE_TOKENS = {
    "#unknown!",
    "#value!",
    "#div/0!",
    "#ref!",
    "#name?",
    "#null!",
    "#num!",
    "#n/a",
}

# أكواد ألوان ANSI للّوغ (بدون مكتبات إضافية)
RESET = "\033[0m"
CYAN = "\033[36m"
YELLOW = "\033[33m"
GREEN = "\033[32m"
RED = "\033[31m"
MAGENTA = "\033[35m"


def log_info(msg: str) -> None:
    print(f"{CYAN}[INFO]{RESET} {msg}")


def log_warn(msg: str) -> None:
    print(f"{YELLOW}[WARN]{RESET} {msg}")


def log_success(msg: str) -> None:
    print(f"{GREEN}[OK]{RESET}   {msg}")


def log_error(msg: str) -> None:
    print(f"{RED}[ERROR]{RESET} {msg}")


def log_debug(msg: str) -> None:
    print(f"{MAGENTA}[DEBUG]{RESET} {msg}")


# ==========================
# إدارة فولدر الصور
# ==========================

def ensure_clean_images_dir() -> None:
    """
    يتأكد من وجود فولدر الصور ويمسح أي ملفات موجودة بداخله قبل البدء.
    يعمل داخل IMAGES_DIR المحددة تحت ROOT_DIR.
    """
    if os.path.isdir(IMAGES_DIR):
        for name in os.listdir(IMAGES_DIR):
            path = os.path.join(IMAGES_DIR, name)
            if os.path.isfile(path):
                os.remove(path)
    else:
        os.makedirs(IMAGES_DIR, exist_ok=True)

    log_info(f"Images directory ready and cleaned: '{IMAGES_DIR}'")



def get_image_bytes(img) -> bytes | None:
    """
    تحاول استخراج بايتات الصورة من كائن openpyxl Image.
    نعتمد على الخاصية _data (قد تكون دالة أو بايتات جاهزة).
    """
    data_attr = getattr(img, "_data", None)

    if callable(data_attr):
        try:
            return data_attr()
        except Exception as e:
            log_warn(f"Failed to call img._data(): {e}")

    if isinstance(data_attr, (bytes, bytearray)):
        return bytes(data_attr)

    # لو فشلنا بكل الطرق
    log_warn("Could not extract image bytes from Image object.")
    return None


def guess_image_ext(data: bytes) -> str:
    """
    تخمين بسيط لامتداد الصورة من أول بايتات بدون استخدام مكتبات إضافية.
    لو ما عرفنا النوع نرجّع 'jpg' افتراضياً.
    """
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"\xff\xd8"):
        return "jpg"
    if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        return "gif"
    if data.startswith(b"BM"):
        return "bmp"
    # افتراضي
    return "jpg"


# ==========================
# حساب إحداثيات Y للصفوف
# ==========================

def compute_row_y_map(ws):
    """
    تحسب إحداثيات Y (بالـ EMU) لكل صف:
    top_y[row] = موضع بداية الصف من الأعلى
    bottom_y[row] = موضع نهاية الصف
    نعتمد على ارتفاع الصفوف (إن كان مخصصاً) أو ارتفاع الديفولت من sheet_format.
    """
    sheet_format = ws.sheet_format
    default_height_points = sheet_format.defaultRowHeight or 15  # نقاط
    EMU_PER_POINT = 12700  # ثابت تحويل من نقاط إلى EMU

    top_y = {}
    bottom_y = {}
    acc = 0

    for r in range(1, ws.max_row + 1):
        h = ws.row_dimensions[r].height
        if h is None:
            h = default_height_points
        top_y[r] = acc
        acc += h * EMU_PER_POINT
        bottom_y[r] = acc

    log_info(
        f"Row Y mapping computed using default height {default_height_points} pt "
        f"(~{default_height_points * EMU_PER_POINT:.0f} EMU per row when not custom)."
    )
    return top_y, bottom_y


# ==========================
# منطق استخراج الباكجات
# ==========================

def find_first_header_row(ws) -> int:
    """
    تبحث عن أول سطر يحتوي على أي من الكلمات:
    Part Number / Description / Qty
    في أي عمود من الأعمدة، وترجع رقم السطر (1-based).
    لو لم تجده ترجع 0.
    """
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell_value = ws.cell(row=row, column=col).value
            if not cell_value:
                continue

            text = str(cell_value).strip().lower()
            if any(keyword in text for keyword in HEADER_KEYWORDS):
                log_info(
                    f"First header row detected at row {row} "
                    f"(col {col}) with value '{cell_value}'"
                )
                return row

    log_warn("No header row found with Part Number / Description / Qty in any column.")
    return 0


def build_packages(ws, top_y, bottom_y) -> List[Dict[str, Any]]:
    """
    تبني قائمة الباكجات اعتماداً على العمود الأول:
    - نبحث أولاً عن سطر الهيدر في أي عمود.
    - نبدأ من السطر السابق للهيدر، ونستخدم العمود الأول فقط لاكتشاف أسماء الباكجات.
    - كل خلية غير فارغة (وليست كلمة هيدر، وليست قيمة خطأ مثل #VALUE!) في العمود الأول
      تعتبر اسم باكج وبداية له.
    - أسطر الباكج تمتد من بداية الباكج حتى السطر السابق لبداية الباكج التالي.
    - نحسب لكل باكج أيضاً y_start و y_end على محور Y.
    """
    header_row = find_first_header_row(ws)
    if header_row <= 1:
        log_error("Cannot determine package start row (header row not found or at first row).")
        return []

    # بداية البحث عن أسماء الباكجات: السطر السابق للهيدر
    start_row = header_row - 1
    log_info(f"Package detection will start from row {start_row} (row above first header).")

    packages: List[Dict[str, Any]] = []
    current_package = None

    for row in range(start_row, ws.max_row + 1):
        # العمود الأول فقط: هو الذي يحتوي أسماء الباكجات
        cell_value = ws.cell(row=row, column=1).value
        if cell_value is None:
            # سطر فارغ في العمود الأول → يبقى ضمن الباكج الحالي إن وجد
            continue

        text = str(cell_value).strip()
        lower_text = text.lower()

        # نتجاهل ظهور كلمات الهيدر في العمود الأول لو حصلت
        if any(keyword in lower_text for keyword in HEADER_KEYWORDS):
            continue

        # نتجاهل قيم الأخطاء مثل #VALUE! بحيث لا تفتح باكج جديدة
        if lower_text in EXCLUDED_PACKAGE_TOKENS:
            log_debug(f"Ignoring error-like value at row {row}: '{text}'")
            continue

        # هنا خلية غير فارغة وليست هيدر وليست قيمة خطأ → اسم باكج جديد
        if current_package is not None:
            # ننهي الباكج السابق عند السطر السابق
            current_package["end_row"] = row - 1
            # نحسب نطاق الـ Y
            current_package["y_start"] = top_y[current_package["start_row"]]
            current_package["y_end"] = bottom_y[current_package["end_row"]]
            packages.append(current_package)

        current_package = {
            "name": text,
            "start_row": row,
            "end_row": None,   # سنحددها لاحقاً
            "y_start": None,
            "y_end": None,
            "images": [],      # صور OneCellAnchor
            "abs_images": [],  # صور AbsoluteAnchor
            "uid": None,       # سيملأ لاحقاً
            "image_filename": None,
            "category": None,  # سنملأها لاحقاً من العمود F
        }
        log_debug(f"New package detected at row {row}: '{text}'")

    # إنهاء آخر باكج إن وجد
    if current_package is not None:
        current_package["end_row"] = ws.max_row
        current_package["y_start"] = top_y[current_package["start_row"]]
        current_package["y_end"] = bottom_y[current_package["end_row"]]
        packages.append(current_package)

    if not packages:
        log_warn("No packages were detected.")
        return []

    log_success(f"{len(packages)} packages detected.")
    return packages



def fill_packages_categories(ws, packages: List[Dict[str, Any]]) -> None:
    """
    لكل باكج:
    - نبحث في العمود F (col=6) ضمن مجال أسطر الباكج [start_row..end_row]
    - نأخذ أول خانة فيها نص (بعد strip)، وهذه هي Category
    - إذا لم يوجد أي نص، تبقى Category = None
    """
    CATEGORY_COL = 6  # العمود F

    for pkg in packages:
        category = None

        for row in range(pkg["start_row"], pkg["end_row"] + 1):
            cell_value = ws.cell(row=row, column=CATEGORY_COL).value
            if isinstance(cell_value, str):
                text = cell_value.strip()
                if text:
                    category = text
                    break

        pkg["category"] = category
        log_debug(
            f"Package '{pkg['name']}' rows [{pkg['start_row']}-{pkg['end_row']}]: "
            f"Category = '{category}'"
        )


# ==========================
# ربط الصور بالباكجات
# ==========================

def find_package_for_row(packages: List[Dict[str, Any]], row: int) -> Dict[str, Any] | None:
    """
    تبحث عن الباكج الذي يحتوي هذا السطر:
    start_row <= row <= end_row
    """
    for pkg in packages:
        if pkg["start_row"] <= row <= pkg["end_row"]:
            return pkg
    return None


def find_package_for_y_center(packages: List[Dict[str, Any]], center_y: int) -> Dict[str, Any] | None:
    """
    تبحث عن الباكج الذي يحتوي مركز الصورة عمودياً:
    y_start <= center_y < y_end
    """
    for pkg in packages:
        if pkg["y_start"] is None or pkg["y_end"] is None:
            continue
        if pkg["y_start"] <= center_y < pkg["y_end"]:
            return pkg
    return None

def link_images_to_packages(ws, packages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    - تمرّ على الصور وتربطها بالباكجات (OneCellAnchor → حسب السطر، AbsoluteAnchor → حسب الـ Y).
    - لكل باكج نولّد UID حتى لو لم يكن لها صورة.
    - لو وُجدت صورة وتم حفظها:
        pkg["uid"] = uid
        pkg["image_filename"] = اسم ملف الصورة
      لو لم تُحفظ:
        pkg["uid"] = uid
        pkg["image_filename"] = None
    - لا تبني صفوف إكسل، فقط تُحدّث كائنات الباكجات.
    """
    images = getattr(ws, "_images", [])
    log_info(f"Total images found: {len(images)}")
    log_info(f"Anchor types count: {Counter(type(img.anchor).__name__ for img in images)}")

    unmatched_images = []

    # --- المرحلة الأولى: ربط الصور بالباكجات (ملء images / abs_images) ---
    for idx, img in enumerate(images):
        anchor = img.anchor
        tname = type(anchor).__name__

        if tname in ["OneCellAnchor", "TwoCellAnchor"] and hasattr(anchor, "_from") and anchor._from is not None:
            # الربط عن طريق الأسطر
            fm = anchor._from
            row_zero_based = getattr(fm, "row", 0)
            row_excel = row_zero_based + 1  # تحويل من 0-based إلى 1-based

            pkg = find_package_for_row(packages, row_excel)
            if pkg:
                # لو كان للباكج صورة سابقة (من أي نوع) نتجاهل هذه الصورة الإضافية
                if pkg["images"] or pkg["abs_images"]:
                    log_warn(
                        f"Ignoring extra image #{idx} for package '{pkg['name']}' "
                        f"(already has an image)."
                    )
                else:
                    pkg["images"].append(idx)
                    log_success(
                        f"Image #{idx} (OneCellAnchor at row {row_excel}) linked to package "
                        f"'{pkg['name']}' [rows {pkg['start_row']} - {pkg['end_row']}]"
                    )
            else:
                unmatched_images.append(idx)
                log_warn(
                    f"Image #{idx} (OneCellAnchor at row {row_excel}) "
                    f"could not be matched to any package."
                )

        elif tname == "AbsoluteAnchor" and hasattr(anchor, "pos") and anchor.pos is not None:
            # الربط عن طريق محور Y (pos.y + cy/2)
            pos = anchor.pos
            ext = getattr(anchor, "ext", None)

            if ext is None:
                unmatched_images.append(idx)
                log_warn(
                    f"Image #{idx} (AbsoluteAnchor) has no ext; skipped for package mapping."
                )
                continue

            y_top = getattr(pos, "y", None)
            cy = getattr(ext, "cy", None)

            if y_top is None or cy is None:
                unmatched_images.append(idx)
                log_warn(
                    f"Image #{idx} (AbsoluteAnchor) missing pos.y or ext.cy; skipped for mapping."
                )
                continue

            center_y = y_top + cy / 2
            pkg = find_package_for_y_center(packages, center_y)

            if pkg:
                # نفس الفكرة: إن كان للباكج صورة مسبقاً نتجاهل هذه
                if pkg["images"] or pkg["abs_images"]:
                    log_warn(
                        f"Ignoring extra image #{idx} for package '{pkg['name']}' "
                        f"(already has an image)."
                    )
                else:
                    pkg["abs_images"].append(idx)
                    log_success(
                        f"Image #{idx} (AbsoluteAnchor center_y={center_y:.0f}) linked to package "
                        f"'{pkg['name']}' [Y {pkg['y_start']:.0f} - {pkg['y_end']:.0f}]"
                    )
            else:
                unmatched_images.append(idx)
                log_warn(
                    f"Image #{idx} (AbsoluteAnchor center_y={center_y:.0f}) "
                    f"could not be matched to any package."
                )

        else:
            unmatched_images.append(idx)
            log_warn(f"Image #{idx} with anchor type '{tname}' could not be processed for mapping.")

    # --- المرحلة الثانية: توليد UID لكل باكج وحفظ الصورة إن وجدت ---
    log_info("=== Package list (id + optional image) ===")
    print("package_name\tstart_row\tid\timage")

    for pkg in packages:
        # نختار أولاً صورة من نوع OneCellAnchor (لو وجدت)، ثم Absolute
        img_idx = None
        if pkg["images"]:
            img_idx = pkg["images"][0]
        elif pkg["abs_images"]:
            img_idx = pkg["abs_images"][0]

        uid = str(uuid4())
        filename = None  # None تعني لا يوجد ملف صورة

        if img_idx is not None:
            img_obj = images[img_idx]
            img_bytes = get_image_bytes(img_obj)

            if not img_bytes:
                log_warn(
                    f"Image bytes for image #{img_idx} (package '{pkg['name']}') "
                    f"could not be extracted."
                )
            else:
                ext = guess_image_ext(img_bytes)
                filename = f"{uid}.{ext}"
                filepath = os.path.join(IMAGES_DIR, filename)

                try:
                    with open(filepath, "wb") as f:
                        f.write(img_bytes)

                    log_success(
                        f"Saved image #{img_idx} for package '{pkg['name']}' "
                        f"as '{filename}'"
                    )
                except Exception as e:
                    log_error(f"Failed to save image for package '{pkg['name']}': {e}")
                    filename = None  # نفشل فنرجعها None

        pkg["uid"] = uid
        pkg["image_filename"] = filename

        print(f"{pkg['name']}\t{pkg['start_row']}\t{uid}\t{filename or ''}")

    if unmatched_images:
        log_warn(f"Unmatched images: {unmatched_images}")

    return packages

# ==========================
# معالجة ملف واحد
# ==========================

def process_workbook(path: str) -> List[tuple[str, str, str, str, str]]:
    """
    تعالج ملف إكسل واحد:
    - تفتح الملف
    - تبني خريطة Y
    - تبني الباكجات
    - تملأ Category من العمود F
    - تربط الصور بالباكجات (وتحفظ الصور بالفولدر وتملأ uid / image_filename)
    - ترجع قائمة صفوف:
      (PackageId, ImagePath, TitleTrim, PackageName, Category)
    """
    basename = os.path.basename(path)
    title_trim = os.path.splitext(basename)[0].strip()  # هذا هو Title - TRIM

    log_info(f"Opening workbook: {basename}")

    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        log_error(f"Failed to open '{basename}': {e}")
        return []

    ws = wb.active

    # حساب خريطة Y للصفوف
    top_y, bottom_y = compute_row_y_map(ws)

    # بناء الباكجات مع y_start / y_end
    packages = build_packages(ws, top_y, bottom_y)
    if not packages:
        log_warn(f"No packages found in '{basename}'. Skipping image mapping.")
        return []

    # أولاً: ملء الفئة Category من العمود F
    fill_packages_categories(ws, packages)

    # ثانياً: ربط الصور + توليد uid + حفظ الصور
    link_images_to_packages(ws, packages)

    # الآن نبني صفوف الإكسل لهذا الملف فقط
    rows_for_excel: List[tuple[str, str, str, str, str]] = []
    for pkg in packages:
        uid = pkg.get("uid")
        image_filename = pkg.get("image_filename") or ""
        pkg_name = pkg["name"]
        category = pkg.get("category") or ""

        rows_for_excel.append((uid, image_filename, title_trim, pkg_name, category))

    return rows_for_excel


# ==========================
# الدالة الرئيسية
# ==========================

def main():
    """
    - يحضر فولدر الصور ويمسح محتوياته.
    - يبحث عن كل ملفات .xlsx في الروت الحالي (باستثناء ملفات إكسل المؤقتة ~$.)
    - يعالج كل ملف على حدة ويجمع الصفوف.
    - يبني ملف إكسل جديد packages_data.xlsx
      يحتوي الأعمدة: id, image, package name
      مع إدراج الصور في عمود image.
    """
    ensure_clean_images_dir()

    files = [
        f for f in os.listdir(ROOT_DIR)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    ]

    if not files:
        log_error(f"No .xlsx files found in ROOT_DIR: {ROOT_DIR}")
        return

    log_info(f"Found {len(files)} .xlsx file(s) in ROOT_DIR: {ROOT_DIR}")

    all_rows: List[tuple[str, str, str]] = []

    for fname in files:
        print()
        print(f"{MAGENTA}========== Processing file: {fname} =========={RESET}")
        full_path = os.path.join(ROOT_DIR, fname)
        rows = process_workbook(full_path)
        all_rows.extend(rows)

    if not all_rows:
        log_warn("No rows were collected. Excel data file will not be created.")
        return

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "packages"

    # عرض الأعمدة الأساسية (عدّلها كما تحب)
    ws_out.column_dimensions["A"].width = 40
    ws_out.column_dimensions["B"].width = 40
    ws_out.column_dimensions["C"].width = 40
    ws_out.column_dimensions["D"].width = 40

    # الهيدر المطلوب
    # الهيدر المطلوب (مع Category قبل delete)
    ws_out.append([
        "PackageId",
        "ImagePath",
        "Title - TRIM",
        "PackageName",
        "No",
        "PartNo",
        "Part Name And Standard",
        "QTY",
        "Category",
        "delete",
        "price",
        "Description",
        "Old Part No.",
        "Names and specifications of old parts",
        "note",
        "is_red",
        "is_line",
        "is_deleted",
        "is_orange",
        "is_pink",
        "is_yellow",
        "internal_notes",
    ])

    # البيانات + إدراج الصور
    for row_idx, (uid, filename, title_trim, pkg_name, category) in enumerate(all_rows, start=2):
        ws_out.append([
            uid,          # PackageId
            filename,     # ImagePath
            title_trim,   # Title - TRIM
            pkg_name,     # PackageName
            "",           # No
            "",           # PartNo
            "",           # Part Name And Standard
            "",           # QTY
            category,     # Category
            "",           # delete
            "",           # price
            "",           # Description
            "",           # Old Part No.
            "",           # Names and specifications of old parts
            "",           # note
            "",           # is_red
            "",           # is_line
            "",           # is_deleted
            "",           # is_orange
            "",           # is_pink
            "",           # is_yellow
            "",           # internal_notes
        ])

        if filename:
            img_path = os.path.join(IMAGES_DIR, filename)
            if os.path.exists(img_path):
                try:
                    xl_img = XLImage(img_path)
                    xl_img.width = 80
                    xl_img.height = 80
                    ws_out.add_image(xl_img, f"B{row_idx}")
                    ws_out.row_dimensions[row_idx].height = 60
                except Exception as e:
                    log_warn(f"Failed to embed image '{img_path}' into Excel: {e}")

    wb_out.save(OUTPUT_EXCEL)
    log_success(f"Data Excel file created: '{OUTPUT_EXCEL}'")


if __name__ == "__main__":
    main()
