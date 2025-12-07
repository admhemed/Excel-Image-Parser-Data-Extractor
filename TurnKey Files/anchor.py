import os
from openpyxl import load_workbook

ROOT_DIR = "."  # عدّلها لو حابب فولدر معيّن

anchor_types = set()


def collect_anchor_type(obj):
    """إضافة نوع الأنكور (مع المسار الكامل) إلى الـ set."""
    if obj is None:
        return
    t = type(obj)
    anchor_types.add(f"{t.__module__}.{t.__name__}")


for dirpath, dirnames, filenames in os.walk(ROOT_DIR):
    for fname in filenames:
        # نتجاهل ملفات الإكسل المؤقتة
        if not fname.lower().endswith((".xlsx", ".xlsm")) or fname.startswith("~$"):
            continue

        fpath = os.path.join(dirpath, fname)
        try:
            wb = load_workbook(fpath, data_only=True)
        except Exception as e:
            print(f"خطأ في فتح الملف {fpath}: {e}")
            continue

        for ws in wb.worksheets:
            # 1) الأنكور المرتبط بالصور نفسها
            for img in getattr(ws, "_images", []):
                collect_anchor_type(getattr(img, "anchor", None))

            # 2) الأنكور داخل الـ drawing (oneCell / twoCell / absolute)
            drawing = getattr(ws, "_drawing", None)
            if drawing is not None:
                # في بعض نسخ openpyxl الكائن الفعلي يكون في الخاصية _drawing
                container = getattr(drawing, "_drawing", drawing)

                for attr_name in ("oneCellAnchor", "twoCellAnchor", "absoluteAnchor"):
                    anchors = getattr(container, attr_name, [])
                    # anchors ممكن تكون list أو كائن واحد
                    if isinstance(anchors, (list, tuple)):
                        for anc in anchors:
                            collect_anchor_type(anc)
                    else:
                        collect_anchor_type(anchors)

# نحولهم لقائمة مرتبة فقط للعرض
anchor_types_list = sorted(anchor_types)

print("أنواع الـ anchor الموجودة في كل ملفات الإكسل (بدون تكرار):")
for t in anchor_types_list:
    print("-", t)
