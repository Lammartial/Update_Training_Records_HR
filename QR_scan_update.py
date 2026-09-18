"""
Incremental QR scan to populate a training evidence Excel template.

- Reads the SUMMARY sheet of an existing populated Excel file.
- Skips PDFs that are already marked as "OK".
- Processes new PDFs or retries PDFs marked as "ERROR".
- Copies an existing training sheet, scrubs it clean, and dynamically resizes it for new data.
- Updates the Excel file in place.
"""

import datetime as dt
import os
from pathlib import Path
from copy import copy
import re

import cv2
import fitz
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

try:
    import zxingcpp  # type: ignore
except Exception:
    zxingcpp = None

try:
    from pyzbar.pyzbar import decode as zbar_decode
except Exception:
    zbar_decode = None


# ========= CONFIG =========
USER_HOME = os.path.expanduser("~")

# Automatically detect which RRC SharePoint folder structure is synced
if os.path.exists(USER_HOME + r"\RRC power solutions\RRC VN - 999_SHARE_VN"):
    BASE_DIR = USER_HOME + r"\RRC power solutions\RRC VN - 999_SHARE_VN"
    print("User synced the 999_SHARE_VN folder separately!")
elif os.path.exists(USER_HOME + r"\RRC power solutions\RRC VN - Documents\999_SHARE_VN"):
    BASE_DIR = USER_HOME + r"\RRC power solutions\RRC VN - Documents\999_SHARE_VN"
    print("User synced the whole Sharepoint library Documents folder!")
else:
    raise FileNotFoundError("Cannot find the RRC 999_SHARE_VN folder.")

# Folder containing training evidence PDFs to scan
# INPUT_DIR = r"C:\Training HR scripts\PDF_to_import"
INPUT_DIR = BASE_DIR + r"\280_HR\004_Training\100_Training\Training_Evidence_Scan\PDF_to_import"

# The EXISTING Excel file that has already been partially processed.
# The script will read its SUMMARY sheet and append new sheets to this file.
# EXISTING_XLSX = r"C:\Training HR scripts\Input\Training_Evidence.xlsx"
EXISTING_XLSX = BASE_DIR + r"\280_HR\004_Training\100_Training\Training_Evidence_Scan\XLSX_to_import\Training_Evidence.xlsx"

# The file to save to (Set to EXISTING_XLSX to update in-place, or a new path to create a copy)
OUTPUT_XLSX = EXISTING_XLSX

# OUTPUT_DIR = r"C:\Training HR scripts\Input"
OUTPUT_DIR = BASE_DIR + r"\280_HR\004_Training\100_Training\Training_Evidence_Scan\XLSX_to_import"

# Optimized Multi-Pass Strategy Settings
SAVE_DEBUG_IMAGES = False   # Set to True to output debug image with QR code bounding box scan results.
DPI = 200
SCAN_DPI_CHOICES = (200,)
SCAN_CLUSTER_TOL = 0.01  # relative page-coordinate tolerance when merging multi-pass detections
# ==========================    

def parse_numeric_text(x):
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    else:
        s = str(x).strip()
        if not s:
            raise ValueError("Empty numeric value")
        s = s.replace(" ", "").replace("\u00a0", "")
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)

def excel_date_to_datetime(x):
    try:
        val = parse_numeric_text(x)
    except (ValueError, TypeError):
        return x
    return dt.datetime(1899, 12, 30) + dt.timedelta(days=val)

def clean_duration(x):
    return int(round(parse_numeric_text(x)))

def is_control_qr(text: str) -> bool:
    t = (text or "").strip()
    return t.startswith("CHK") and "ALG=" in t and "VAL=" in t

def parse_control_qr_expected_count(text: str):
    t = (text or "").strip()
    if not is_control_qr(t):
        return None
    parts = [p.strip() for p in t.split("|") if p.strip()]
    for part in parts:
        if part.upper().startswith("VAL="):
            val = part.split("=", 1)[1].strip()
            try:
                return int(float(val))
            except (ValueError, TypeError):
                raise ValueError(f"Invalid control QR VAL value: {val!r} in {t!r}")
    raise ValueError(f"Control QR missing VAL field: {t!r}")

def is_numeric_id(text: str) -> bool:
    t = (text or "").strip()
    return bool(t) and all(ch.isdigit() for ch in t)

def pdf_pages_to_images(pdf_path: str, dpi: int = 300):
        # Disable anti-aliasing globally before rendering to prevent muddy module boundaries
    try:
        fitz.set_aa_level(0)
    except Exception:
        pass
    doc = fitz.open(pdf_path)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pages = []
    for p in range(len(doc)):
        page = doc[p]
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
        pages.append((p + 1, img))
    doc.close()
    return pages

def _qr_formats():
    if zxingcpp is None:
        return None
    BFmts = getattr(zxingcpp, "BarcodeFormats", None)
    if BFmts is None:
        return None
    for name in ("QRCode", "QR_CODE", "QRCODE", "QrCode"):
        val = getattr(BFmts, name, None)
        if val is not None:
            return val
    return getattr(BFmts, "Any", None)

QR_FMTS = _qr_formats()

def _position_to_quad(pos):
    if pos is None:
        return None
    pts = None
    try:
        pts = list(pos)
    except TypeError:
        pts = None
    if (pts is None) and hasattr(pos, "points"):
        try:
            pts = list(pos.points)
        except Exception:
            pts = None
    if pts is None:
        corner_names = [
            ("top_left", "topLeft"),
            ("top_right", "topRight"),
            ("bottom_right", "bottomRight"),
            ("bottom_left", "bottomLeft"),
        ]
        corners = []
        for names in corner_names:
            found = None
            for n in names:
                if hasattr(pos, n):
                    found = getattr(pos, n)
                    break
            if found is None:
                for n in [names[0] + "_", names[1] + "_"]:
                    if hasattr(pos, n):
                        found = getattr(pos, n)
                        break
            if found is not None:
                corners.append(found)
        if len(corners) == 4:
            pts = corners
    if not pts:
        return None
    quad = []
    for p in pts[:4]:
        if isinstance(p, (list, tuple)) and len(p) >= 2:
            x, y = float(p[0]), float(p[1])
        elif hasattr(p, "x") and hasattr(p, "y"):
            x, y = float(p.x), float(p.y)
        else:
            continue
        quad.append([x, y])
    if len(quad) < 4:
        return None
    return np.array(quad[:4], dtype=float)

def decode_qr_zxing(img_bgr):
    if zxingcpp is None:
        return []
    img_rgb = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)
    kwargs = {
        "text_mode": getattr(zxingcpp.TextMode, "UTF8", zxingcpp.TextMode.HRI),
        "try_rotate": True,
        "try_downscale": True,
        "return_errors": False,
    }
    if QR_FMTS is not None:
        kwargs["formats"] = QR_FMTS

    results = zxingcpp.read_barcodes(img_rgb, **kwargs)

    out = []
    for r in results:
        txt = (r.text or "").strip()
        if not txt:
            continue
        quad = _position_to_quad(getattr(r, "position", None))
        if quad is None and hasattr(r, "bounds") and r.bounds:
            b = r.bounds
            x, y, w, h = b.x, b.y, b.width, b.height
            quad = np.array([[x, y], [x + w, y], [x + w, y + h], [x, y + h]], dtype=float)
        if quad is None:
            continue
        xs, ys = quad[:, 0], quad[:, 1]
        w = float(xs.max() - xs.min())
        h = float(ys.max() - ys.min())
        
        # GEOMETRIC FILTER: Real QR codes are squares. 
        # If the aspect ratio isn't between 0.7 and 1.4, it's likely text.
        if h == 0 or (w / h < 0.7) or (w / h > 1.4):
            continue

        out.append({
            "text": txt,
            "bbox": quad,
            "left": float(xs.min()),
            "top": float(ys.min()),
            "width": w,
            "height": h,
            "cx": float(xs.mean()),
            "cy": float(ys.mean()),
        })
    return out

def decode_qr_pyzbar(img_bgr):
    """
    Decode QR codes with pyzbar when available.

    If pyzbar is not installed, return an empty list instead of raising.
    This lets blank pages be skipped cleanly.
    """
    if zbar_decode is None:
        return []
    
    results = zbar_decode(img_bgr)

    out = []
    for r in results:
        # Fallback check: Ignore anything pyzbar finds that isn't a QRCODE
        if getattr(r, "type", "") != "QRCODE":
            continue

        txt = r.data.decode("utf-8", errors="replace").strip()
        if not txt:
            continue

        rect = r.rect
        w = float(rect.width)
        h = float(rect.height)
        
        # GEOMETRIC FILTER: Apply the same squareness check here
        if h == 0 or (w / h < 0.7) or (w / h > 1.4):
            continue

        out.append({
            "text": txt,
            "bbox": None,
            "left": float(rect.left),
            "top": float(rect.top),
            "width": w,
            "height": h,
            "cx": float(rect.left + w / 2),
            "cy": float(rect.top + h / 2),
        })
    return out

def detect_page_qrs(img_bgr):
    """
    Scan the page comprehensively using multiple engine variations and image 
    preprocessing routines, ensuring high-density codes are captured alongside sparse ones.
    """
    all_dets = []

    def merge_detections(new_dets):
        for nd in new_dets:
            is_dup = False
            # Deduplicate by spatial proximity on the page pass (within 25 pixels)
            for d in all_dets:
                if abs(nd["cx"] - d["cx"]) < 25 and abs(nd["cy"] - d["cy"]) < 25:
                    is_dup = True
                    if not d["text"] and nd["text"]:
                        d["text"] = nd["text"]
                    break
            if not is_dup:
                all_dets.append(nd)

    # Strategy A: Raw Matrix Scans
    merge_detections(decode_qr_zxing(img_bgr))
    merge_detections(decode_qr_pyzbar(img_bgr))

    # Strategy B: High-Pass Sharpening Filter
    kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]], dtype=np.float32)
    sharpened = cv2.filter2D(img_bgr, -1, kernel)
    merge_detections(decode_qr_zxing(sharpened))
    merge_detections(decode_qr_pyzbar(sharpened))

    # Strategy C: Adaptive Gaussian Thresholding (Crucial for tight, high-density grids)
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    adaptive_thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 15, 5    #11, 15, 21, 31
    )
    adaptive_bgr = cv2.cvtColor(adaptive_thresh, cv2.COLOR_GRAY2BGR)
    merge_detections(decode_qr_zxing(adaptive_bgr))
    merge_detections(decode_qr_pyzbar(adaptive_bgr))

    # Strategy D: Global Otsu Thresholding
    _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    otsu_bgr = cv2.cvtColor(otsu, cv2.COLOR_GRAY2BGR)
    merge_detections(decode_qr_zxing(otsu_bgr))
    merge_detections(decode_qr_pyzbar(otsu_bgr))
    if not all_dets:
        return []
    # ========================================================
    # NEW LOGIC: Filter Outliers and Overlapping Hallucinations
    # ========================================================
    
    # 1. Size Outlier Filter
    widths = [d["width"] for d in all_dets]
    heights = [d["height"] for d in all_dets]
    median_w = np.median(widths)
    median_h = np.median(heights)

    size_filtered = []
    for d in all_dets:
        # Give the Control QR a "VIP pass" to bypass size filtering entirely
        if is_control_qr(d.get("text", "")):
            size_filtered.append(d)
            continue

        # For all normal QRs: Drop boxes <30% smaller than the median. 
        # if d["width"] <= median_w * 1.4:
        if d["width"] >= median_w * 0.3 and d["height"] >= median_h * 0.3:
            size_filtered.append(d)

    # 2. Overlap/Contained (QR within a QR) Filter
    # Sort boxes: prioritize those that actually decoded text, then by area size.
    size_filtered.sort(key=lambda d: (bool(d.get("text")), d["width"] * d["height"]), reverse=True)

    final_dets = []
    for d in size_filtered:
        is_contained = False
        for kept in final_dets:
            x_left = max(d["left"], kept["left"])
            y_top = max(d["top"], kept["top"])
            x_right = min(d["left"] + d["width"], kept["left"] + kept["width"])
            y_bottom = min(d["top"] + d["height"], kept["top"] + kept["height"])

            if x_right > x_left and y_bottom > y_top:
                intersection_area = (x_right - x_left) * (y_bottom - y_top)
                min_area = min(d["width"] * d["height"], kept["width"] * kept["height"])
                
                # If more than 50% of the smaller box overlaps the larger one, it's a hallucination
                if min_area > 0 and (intersection_area / min_area) > 0.5:
                    is_contained = True
                    break
                    
        if not is_contained:
            final_dets.append(d)

    return final_dets


def order_qrs_top_bottom_left(qrs):
    if not qrs: return qrs
    qrs = sorted(qrs, key=lambda r: (r["cy"], r["cx"]))
    heights = [r["height"] for r in qrs if r["height"] > 0]
    line_tol = (np.median(heights) * 0.6) if heights else 20.0
    groups, row = [], [qrs[0]]
    for r in qrs[1:]:
        if abs(r["cy"] - row[-1]["cy"]) <= line_tol:
            row.append(r)
        else:
            groups.append(sorted(row, key=lambda x: x["cx"]))
            row = [r]
    groups.append(sorted(row, key=lambda x: x["cx"]))
    return [item for g in groups for item in g]

def order_clusters_top_bottom_left(clusters):
    """Robust row grouping structure for normalized cluster coordinates."""
    if not clusters:
        return clusters
    clusters = sorted(clusters, key=lambda c: (c["ry"], c["rx"]))
    heights = [c["rh"] for c in clusters if c["rh"] > 0]
    line_tol = (np.median(heights) * 0.6) if heights else 0.02
    groups, row = [], [clusters[0]]
    for c in clusters[1:]:
        if abs(c["ry"] - row[-1]["ry"]) <= line_tol:
            row.append(c)
        else:
            groups.append(sorted(row, key=lambda x: x["rx"]))
            row = [c]
    groups.append(sorted(row, key=lambda x: x["rx"]))
    return [item for g in groups for item in g]
def draw_debug(image_bgr, ordered_qrs, page_no, out_path):
    dbg = image_bgr.copy()
    img_h, img_w = dbg.shape[:2]
    for idx, r in enumerate(ordered_qrs, start=1):
        if "rx" in r and "ry" in r:
            w = int(round(float(r.get("rw", 0.0)) * img_w))
            h = int(round(float(r.get("rh", 0.0)) * img_h))
             # Convert center coordinates back to the correct top-left corner
            x = int(round(float(r["rx"]) * img_w - w / 2))
            y = int(round(float(r["ry"]) * img_h - h / 2))

        else:
            x, y = int(r.get("left", 0)), int(r.get("top", 0))
            w, h = int(r.get("width", 0)), int(r.get("height", 0))
        cv2.rectangle(dbg, (x, y), (x + w, y + h), (0, 0, 255), 3)
        cv2.putText(dbg, f"{idx}", (x, max(0, y - 10)),
                    cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 255), 2, cv2.LINE_AA)
        cv2.putText(dbg, r["text"][:18], (x, y + h + 30),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2, cv2.LINE_AA)
    cv2.putText(dbg, f"Page {page_no}", (20, 50),
                cv2.FONT_HERSHEY_SIMPLEX, 1.4, (255, 0, 0), 3, cv2.LINE_AA)
    ok = cv2.imwrite(out_path, dbg)
    if not ok:
        raise PermissionError(f"Failed to write annotated image: {out_path}")

def extract_all_qr_fields(pdf_path: str, dpi: int = 300, save_debug: bool = False, output_dir: str | None = None):
    def _collect_page_candidates(pdf_path_inner: str):
        candidates_by_page = {}
        debug_pages = None
        for pass_dpi in SCAN_DPI_CHOICES:
            if pass_dpi is not None:    
                pages = pdf_pages_to_images(pdf_path_inner, dpi=pass_dpi)
                if debug_pages is None: debug_pages = pages
                for page_no, img_rgb in pages:
                    img_bgr = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR)
                    dets = detect_page_qrs(img_bgr)
                    ordered = order_qrs_top_bottom_left(dets)
                    img_h, img_w = img_rgb.shape[:2]
                    page_bucket = candidates_by_page.setdefault(page_no, [])
                    for d in ordered:
                        dd = dict(d)
                        dd["dpi"] = pass_dpi
                        dd["page_no"] = page_no
                        dd["page_w"] = img_w
                        dd["page_h"] = img_h
                        dd["rx"] = float(dd["cx"]) / float(img_w)
                        dd["ry"] = float(dd["cy"]) / float(img_h)
                        dd["rw"] = float(dd["width"]) / float(img_w)
                        dd["rh"] = float(dd["height"]) / float(img_h)
                        page_bucket.append(dd)
        return candidates_by_page, debug_pages

    def _cluster_page_detections(dets, tol=SCAN_CLUSTER_TOL):
        if not dets: return []
        dets = sorted(dets, key=lambda d: (d["ry"], d["rx"], d["dpi"]))
        clusters = []
        for d in dets:
            match = None
            for c in clusters:
                if abs(d["rx"] - c["rx"]) <= tol and abs(d["ry"] - c["ry"]) <= tol:
                    match = c
                    break
            if match is None:
                clusters.append({
                    "text": d["text"],
                    "rx": d["rx"],
                    "ry": d["ry"],
                    "rw": d["rw"],
                    "rh": d["rh"],
                    "page_no": d["page_no"],
                    "members": [d],
                })
            else:
                match["members"].append(d)
                n = len(match["members"])
                match["rx"] = sum(m["rx"] for m in match["members"]) / n
                match["ry"] = sum(m["ry"] for m in match["members"]) / n
                match["rw"] = sum(m["rw"] for m in match["members"]) / n
                match["rh"] = sum(m["rh"] for m in match["members"]) / n
                if not match["text"] and d["text"]:
                    match["text"] = d["text"]
        return clusters

    candidates_by_page, debug_pages = _collect_page_candidates(pdf_path)
    clustered_pages = {}
    for page_no, dets in candidates_by_page.items():
        raw_clusters = _cluster_page_detections(dets)
        # Sort the final combined page clusters using line-grouped layout logic
        clustered_pages[page_no] = order_clusters_top_bottom_left(raw_clusters)

    if save_debug and output_dir and debug_pages:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        for page_no, img_rgb in debug_pages:
            if page_no not in clustered_pages: continue
            img_bgr = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR)
            dbg_path = str(Path(output_dir) / f"{Path(pdf_path).stem}_p{page_no:03d}_annotated.png")
            draw_debug(img_bgr, clustered_pages[page_no], page_no, dbg_path)

    all_clusters = []
    for page_no in sorted(clustered_pages):
        all_clusters.extend(clustered_pages[page_no])

    raw_all_texts = [c["text"] for c in all_clusters]
    control_qrs = [c["text"] for c in all_clusters if is_control_qr(c["text"])]
    non_control_qrs = [c for c in all_clusters if c["text"].strip() and not is_control_qr(c["text"])]

    if not control_qrs:
        raise ValueError("Error detecting QR codes: no control QR found in the PDF.")

    expected_counts = [parse_control_qr_expected_count(cqr) for cqr in control_qrs]
    if len(set(expected_counts)) != 1:
        raise ValueError(f"Error detecting QR codes: conflicting control QR values found: {control_qrs}")

    expected_non_control_count = expected_counts[0]
    actual_non_control_count = len(non_control_qrs)

    if actual_non_control_count != expected_non_control_count:
        raise ValueError(
            "Error detecting QR codes: expected "
            f"{expected_non_control_count} QR codes excluding the control QR, but detected {actual_non_control_count}."
        )

    if len(non_control_qrs) < 8:
        raise ValueError(f"Expected at least 8 non-control QR codes (7 header + trainer). Found {len(non_control_qrs)}.")

    top_fields = [d["text"] for d in non_control_qrs[:7]]
    trainee_ids = [d["text"] for d in non_control_qrs[7:-1] if d["text"].strip()]
    trainer_employee_id = non_control_qrs[-1]["text"]

    return {
        "top_fields": top_fields,
        "trainee_ids": trainee_ids,
        "trainer_employee_id": trainer_employee_id,
        "all_texts": raw_all_texts,
        "control_qrs": control_qrs,
        "expected_non_control_count": expected_non_control_count,
        "actual_non_control_count": actual_non_control_count,
        "per_page": [(pn, clustered_pages[pn]) for pn in sorted(clustered_pages)],
    }

def resize_trainee_section(ws, start_row, end_marker_row, desired_total_rows, style_source_row=30):
    """
    Resize the trainee section so it contains exactly desired_total_rows rows,
    where the last row is kept blank.
    """    
    current_total_rows = end_marker_row - start_row
    if desired_total_rows < 1:
        desired_total_rows = 1
    # Shrink: delete surplus rows immediately before the next section.
    if desired_total_rows < current_total_rows:
        delete_count = current_total_rows - desired_total_rows
        ws.delete_rows(start_row + desired_total_rows, delete_count)
    # Expand: insert rows immediately before the next section.
    elif desired_total_rows > current_total_rows:
        insert_count = desired_total_rows - current_total_rows
        ws.insert_rows(end_marker_row, insert_count)
        # Copy basic row styling into inserted rows.
        for row in range(end_marker_row, end_marker_row + insert_count):
            copy_row_style(ws, style_source_row, row)
    return start_row + desired_total_rows
def sanitize_sheet_title(raw_title: str) -> str:
    title = re.sub(r'[\[\]\*:/\\\?]', '_', str(raw_title).strip())
    title = re.sub(r'\s+', ' ', title).strip()
    return title[:31] if title else "Training"

def unique_sheet_title(wb, desired_title: str) -> str:
    base = sanitize_sheet_title(desired_title)
    if base not in wb.sheetnames:
        return base
    i = 2
    while True:
        suffix = f"_{i}"
        candidate = base[: max(0, 31 - len(suffix))] + suffix
        if candidate not in wb.sheetnames:
            return candidate
        i += 1

def find_row_containing(ws, needles, search_col=1):
    if isinstance(needles, str): needles = [needles]
    for row in range(1, ws.max_row + 1):
        value = ws.cell(row, search_col).value
        if value is None: continue
        s = str(value)
        if any(n.lower() in s.lower() for n in needles):
            return row
    return None

def clear_row_values(ws, row, start_col=1, end_col=5):
    for col in range(start_col, end_col + 1):
        ws.cell(row, col).value = None

def copy_row_style(ws, source_row, target_row, start_col=1, end_col=5):
    for col in range(start_col, end_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style: dst._style = copy(src._style)
        if src.number_format: dst.number_format = src.number_format
        if src.font: dst.font = copy(src.font)
        if src.fill: dst.fill = copy(src.fill)
        if src.border: dst.border = copy(src.border)
        if src.alignment: dst.alignment = copy(src.alignment)
        if src.protection: dst.protection = copy(src.protection)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    ws.row_dimensions[target_row].hidden = ws.row_dimensions[source_row].hidden


def write_to_template_sheet(ws, extracted: dict, source_pdf: str, dpi: int):
    top_fields = extracted["top_fields"]
    trainee_ids = extracted["trainee_ids"]
    trainer_employee_id = extracted["trainer_employee_id"]

    placeholder_map = {
        "C1": top_fields[0],
        "C3": top_fields[1],
        "C5": top_fields[2],
        "C7": top_fields[3],
        "C9": excel_date_to_datetime(top_fields[4]),
        "C11": excel_date_to_datetime(top_fields[5]),
        "C13": clean_duration(top_fields[6]),
        "E15": trainer_employee_id,
    }

    for cell_ref, value in placeholder_map.items():
        ws[cell_ref] = value

    start_row = 21
    next_section_row = find_row_containing(ws, ["Hiệu quả", "The effectiveness"])
    if next_section_row is None:
        next_section_row = 31

    desired_total_rows = len(trainee_ids) + 1
    section_end_row = resize_trainee_section(
        ws=ws,
        start_row=start_row,
        end_marker_row=next_section_row,
        desired_total_rows=desired_total_rows,
        style_source_row=max(start_row, next_section_row - 1),
    )

    # Populate rows and wipe any legacy data if we copied an already-populated sheet
    for idx, trainee_id in enumerate(trainee_ids):
        row = start_row + idx
        ws[f"A{row}"] = idx + 1
        ws[f"B{row}"] = trainee_id
        # Inject the VLOOKUP formula dynamically to avoid openpyxl link corruption
        # formula_str = f"""=IFERROR(VLOOKUP(VALUE(B{row}),'https://rrcpowersolutions.sharepoint.com/sites/RRCVN/Shared Documents/280_HR/104_ Employee_Management_Table/02. Employee management/[RRC_Employee_Master_Data_2025-07-10.xlsx]MASTER_DATA'!$A$2:$B$1048576,2,0),"")"""
        # ws[f"C{row}"] = formula_str
        # ws[f"C{row}"] = None
        #   
        ws[f"D{row}"] = None
        ws[f"E{row}"] = None

    blank_row = start_row + len(trainee_ids)
    clear_row_values(ws, blank_row, 1, 5)

    for row in range(start_row, blank_row + 1):
        if row > 30 and row < section_end_row:
            copy_row_style(ws, 30, row)

    for cell_ref in ("C9", "C11"):
        if isinstance(ws[cell_ref].value, (dt.datetime, dt.date)):
            ws[cell_ref].number_format = "dd/mm/yyyy hh:mm"
    ws["C13"].number_format = "0"

    return {
        "training_no": top_fields[0],
        "trainee_count": len(trainee_ids),
        "trainer_employee_id": trainer_employee_id,
        "expected_non_control_count": extracted.get("expected_non_control_count", ""),
        "actual_non_control_count": extracted.get("actual_non_control_count", ""),
        "total_qr_detected": len(extracted.get("all_texts", [])),
        "control_qr_detected": len(extracted.get("control_qrs", [])),
    }

def process_pdf_into_new_sheet(wb, template_ws, pdf_path, dpi: int, save_debug: bool, output_dir: str):
    extracted = extract_all_qr_fields(pdf_path, dpi=dpi, save_debug=save_debug, output_dir=output_dir)
    desired_name = f"Training_{pdf_path.name}"
    sheet_name = unique_sheet_title(wb, desired_name)

    ws = wb.copy_worksheet(template_ws)
    ws.title = sheet_name
    meta = write_to_template_sheet(ws, extracted, str(pdf_path), dpi)

    return ws, sheet_name, extracted, meta

def build_summary_sheet(wb, summary_rows):
    if "SUMMARY" in wb.sheetnames:
        del wb["SUMMARY"]
    ws = wb.create_sheet("SUMMARY", 0)
    headers = [
        "PDF_File",
        "Sheet_Name",
        "Training_No.",
        "Trainee_Count",
        "Expected_Non_Control_QR",
        "Actual_Non_Control_QR",
        "Control_QR_Count",
        "Status",
        "Error",
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col).value = header

    for row_idx, item in enumerate(summary_rows, start=2):
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row_idx, col_idx).value = item.get(header, "")

    for i, width in enumerate([38, 26, 18, 14, 24, 22, 18, 16, 60], start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"

def main():
    if not os.path.exists(INPUT_DIR):
        raise FileNotFoundError(f"Input folder not found:\n{INPUT_DIR}")
    if not os.path.exists(EXISTING_XLSX):
        raise FileNotFoundError(f"Existing Excel file not found:\n{EXISTING_XLSX}")

    pdf_paths = sorted([p for p in Path(INPUT_DIR).iterdir() if p.is_file() and p.suffix.lower() == ".pdf"])
    if not pdf_paths:
        print(f"No PDF files found in:\n{INPUT_DIR}")
        return

    print(f"Loading existing workbook: {Path(EXISTING_XLSX).name}")
    wb = load_workbook(EXISTING_XLSX)

    # 1. Read existing SUMMARY into a dictionary-like list to track history
    summary_rows = []
    if "SUMMARY" in wb.sheetnames:
        ws_summary = wb["SUMMARY"]
        headers = [cell.value for cell in ws_summary[1]]
        if headers and "PDF_File" in headers:
            for row in ws_summary.iter_rows(min_row=2, values_only=True):
                if not row[0]: continue
                summary_rows.append(dict(zip(headers, row)))

    # 2. Find a template sheet. (If the blank 'Sheet1' was deleted, we just grab any existing Training sheet and scrub it)
    template_ws = None
    for name in ["Sheet1", "Template"]:
        if name in wb.sheetnames:
            print("Template exists in workbook!")
            template_ws = wb[name]
            print(template_ws)
            break
    if not template_ws:
        for name in wb.sheetnames:
            if name != "SUMMARY":
                template_ws = wb[name]
                break
    if not template_ws:
        raise RuntimeError("Could not find any sheet to use as a template in the existing workbook.")

    created_or_updated = False

    # 3. Process PDFs based on history
    error_pdfs = []
    for pdf_path in pdf_paths:
        # Check if the PDF has been processed before
        existing_idx = next((i for i, r in enumerate(summary_rows) if r.get("PDF_File") == pdf_path.name), None)
        
        if existing_idx is not None:
            if summary_rows[existing_idx].get("Status") == "OK":
                print(f"Skipping {pdf_path.name} -> Already marked OK.")
                continue
            else:
                print(f"Retrying {pdf_path.name} -> Previous Status: {summary_rows[existing_idx].get('Status')}")
        else:
            print(f"Processing new file: {pdf_path.name}")

        try:
            # If we are retrying a file, delete the old corrupted sheet to avoid clutter
            desired_name = f"Training_{pdf_path.name}"
            base_title = sanitize_sheet_title(desired_name)
            
            # Find and remove any sheet that matches this PDF, as long as it's not the one we are using as a template!
            sheets_to_remove = [s for s in wb.sheetnames if s.startswith(base_title)]
            for s in sheets_to_remove:
                if s != template_ws.title:
                    del wb[s]

            # Generate the new sheet
            ws, final_sheet_name, extracted, meta = process_pdf_into_new_sheet(
                wb=wb,
                template_ws=template_ws,
                pdf_path=pdf_path,
                dpi=DPI,
                save_debug=SAVE_DEBUG_IMAGES,
                output_dir=OUTPUT_DIR,
            )
            created_or_updated = True
            
            new_record = {
                "PDF_File": pdf_path.name,
                "Sheet_Name": final_sheet_name,
                "Training_No.": meta["training_no"],
                "Trainee_Count": meta["trainee_count"],
                "Expected_Non_Control_QR": meta["expected_non_control_count"],
                "Actual_Non_Control_QR": meta["actual_non_control_count"],
                "Control_QR_Count": meta["control_qr_detected"],
                "Status": "OK",
                "Error": "",
            }

            print(f"Processed: {pdf_path.name} -> {final_sheet_name}")
            print(f"  Training_No.: {meta['training_no']}")
            print(f"  Trainees: {meta['trainee_count']}")
            print(f"  Total QR detections (including control): {meta['total_qr_detected']}")

            if existing_idx is not None:
                summary_rows[existing_idx] = new_record
            else:
                summary_rows.append(new_record)
                
            print(f"  -> Success: {final_sheet_name} (Trainees: {meta['trainee_count']})")

        except Exception as e:
            error_pdfs.append(pdf_path.name)
            new_record = {
                "PDF_File": pdf_path.name,
                "Sheet_Name": "",
                "Training_No.": "",
                "Trainee_Count": "",
                "Expected_Non_Control_QR": "",
                "Actual_Non_Control_QR": "",
                "Control_QR_Count": "",
                "Status": "ERROR",
                "Error": str(e),
            }
            if existing_idx is not None:
                summary_rows[existing_idx] = new_record
            else:
                summary_rows.append(new_record)
            print(f"  -> ERROR: {e}")

    # 4. Save updates
    if created_or_updated:
        build_summary_sheet(wb, summary_rows)
        wb.save(OUTPUT_XLSX)
        print(f"\nSaved updated workbook to: {OUTPUT_XLSX}")

        if SAVE_DEBUG_IMAGES:
            print(f"Annotated page images saved in: {OUTPUT_DIR}")

        if error_pdfs != []:
            print(f"Batch completed with errors in: {', '.join(error_pdfs)}")
        else:
            print("Batch completed successfully with no errors.")
    else:
        print("\nNo new files or retries needed. Workbook remains unchanged.")
    
if __name__ == "__main__":
    main()