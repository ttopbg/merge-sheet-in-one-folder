import pandas as pd
import re
import unicodedata
from io import BytesIO

# ─── Ánh xạ tên cột chuẩn ───────────────────────────────────────────────────
COLUMN_ALIASES = {
    "lop": [
        "lớp", "lop", "lớp học", "lop hoc", "class",
        "khối lớp", "khoi lop", "lớp/class",
    ],
    "ho_ten": [
        "họ tên", "ho ten", "họ và tên", "ho va ten",
        "tên", "ten", "full name", "name",
        "họ tên học sinh", "ho ten hoc sinh",
        "tên học sinh", "ten hoc sinh", "họ tên hs",
    ],
    "ngay_sinh": [
        "ngày sinh", "ngay sinh",
        "ngày tháng năm sinh", "ngay thang nam sinh",
        "dob", "date of birth", "năm sinh", "nam sinh",
        "ngày/tháng/năm sinh",
    ],
    "gioi_tinh": [
        "giới tính", "gioi tinh", "gender", "sex",
        "gt", "phái", "phai",
    ],
}

STANDARD_NAMES = {
    "ho_ten":    "Họ và tên",
    "lop":       "Lớp",
    "gioi_tinh": "Giới tính",
    "ngay_sinh": "Ngày sinh",
}

PRIORITY_COLS = ["Lớp", "Họ và tên", "Giới tính", "Ngày sinh"]
REQUIRED_COLS = {"ho_ten"}  # Chỉ cần cột Họ và tên; các cột khác có thì lấy, không có để trống

# Số dòng đầu mỗi sheet dùng để dò dòng tiêu đề bảng.
# Phần tiêu đề trang (UBND…, tên trường, "DANH SÁCH HỌC SINH LỚP…", "Năm học…")
# nằm phía trên sẽ tự bị bỏ qua.
HEADER_SCAN_ROWS = 30

# Dòng cuối bảng (không phải học sinh) — bỏ qua nếu ô Họ tên bắt đầu bằng các cụm này
FOOTER_PREFIXES = (
    "tong", "cong", "nguoi lap", "hieu truong", "giao vien",
    "ky ten", "xac nhan", "ghi chu",
)


def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFC", str(text)).strip().lower()
    replacements = [
        (r"[àáâãäåạảấầẩẫậắằẳẵặă]", "a"),
        (r"[èéêëẹẻẽếềểễệ]",        "e"),
        (r"[ìíîïịỉĩ]",              "i"),
        (r"[òóôõöọỏốồổỗộớờởỡợơ]",  "o"),
        (r"[ùúûüụủũứừửữựư]",        "u"),
        (r"[ỳýỹỵỷ]",                "y"),
        (r"[đ]",                    "d"),
    ]
    for pat, rep in replacements:
        text = re.sub(pat, rep, text)
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


_ALIASES_NORM = {k: [_normalize(a) for a in v] for k, v in COLUMN_ALIASES.items()}


def _detect_header_row(df_raw: pd.DataFrame) -> int | None:
    """Quét HEADER_SCAN_ROWS dòng đầu, chọn dòng có cột Họ tên và khớp nhiều cột nhất.
    Nhờ vậy phần tiêu đề trang phía trên bảng được bỏ qua."""
    best_idx, best_score = None, 0
    scan_limit = min(HEADER_SCAN_ROWS, len(df_raw))
    for i in range(scan_limit):
        row_vals = [_normalize(c) for c in df_raw.iloc[i].values if pd.notna(c)]
        found = {
            key for key, aliases in _ALIASES_NORM.items()
            if any(v in aliases for v in row_vals)
        }
        if not REQUIRED_COLS.issubset(found):
            continue
        if len(found) > best_score:
            best_idx, best_score = i, len(found)
    return best_idx


def _map_columns(header_row: pd.Series) -> dict:
    mapping = {}
    for col_idx, cell in enumerate(header_row):
        cell_norm = _normalize(cell)
        for std_key, aliases in _ALIASES_NORM.items():
            if cell_norm in aliases and std_key not in mapping:
                mapping[std_key] = col_idx
    return mapping


def _format_date(val) -> str:
    if not isinstance(val, str) and pd.isnull(val):
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    if s.lower() in ("", "nan", "none", "nat"):
        return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return pd.to_datetime(s, format=fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return s


def _is_student_name(val: str) -> bool:
    if not val or val.lower() in ("nan", "none"):
        return False
    n = _normalize(val)
    if not re.search(r"[a-z]", n):          # toàn số / ký hiệu
        return False
    if n.startswith(FOOTER_PREFIXES):        # dòng tổng / chữ ký cuối bảng
        return False
    return True


def extract_sheet(df_raw: pd.DataFrame, file_name: str) -> tuple[pd.DataFrame | None, int | None]:
    header_idx = _detect_header_row(df_raw)
    if header_idx is None:
        return None, None

    header_row = df_raw.iloc[header_idx]
    col_map = _map_columns(header_row)
    if not REQUIRED_COLS.issubset(col_map.keys()):
        return None, None

    original_col_names = list(header_row.values)
    data_rows = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
    mapped_indices = set(col_map.values())

    # Lớp = tên file đầu vào (bỏ phần đuôi)
    lop_val = file_name.rsplit(".", 1)[0]

    def _get(row, key):
        if key not in col_map:
            return ""
        v = row.iloc[col_map[key]]
        return "" if (not isinstance(v, str) and pd.isnull(v)) else str(v).strip()

    records = []
    for _, row in data_rows.iterrows():
        ho_ten_val = _get(row, "ho_ten")
        if not _is_student_name(ho_ten_val):
            continue

        ngay_sinh_raw = row.iloc[col_map["ngay_sinh"]] if "ngay_sinh" in col_map else ""

        record = {
            STANDARD_NAMES["lop"]:       lop_val,
            STANDARD_NAMES["ho_ten"]:    ho_ten_val,
            STANDARD_NAMES["gioi_tinh"]: _get(row, "gioi_tinh"),
            STANDARD_NAMES["ngay_sinh"]: _format_date(ngay_sinh_raw),
        }
        # Nếu file có sẵn cột Lớp thì giữ lại dưới tên "Lớp_gốc"
        if "lop" in col_map:
            record["Lớp_gốc"] = _get(row, "lop")

        # Giữ các cột gốc khác
        for ci, orig_name in enumerate(original_col_names):
            if ci in mapped_indices:
                continue
            col_label = str(orig_name).strip()
            if col_label.lower() in ("nan", "", "none"):
                continue  # cột không có tiêu đề (thường là cột trống) — bỏ
            if col_label in PRIORITY_COLS or col_label == "Lớp_gốc":
                col_label = f"{col_label}_gốc"
            record[col_label] = row.iloc[ci]

        records.append(record)

    return (pd.DataFrame(records) if records else None), header_idx


def merge_excel_files(uploaded_files: list) -> tuple[pd.DataFrame, list[str]]:
    logs = []
    frames = []

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        try:
            if hasattr(uploaded_file, "seek"):
                uploaded_file.seek(0)
            raw_bytes = uploaded_file.read()
            ext = file_name.rsplit(".", 1)[-1].lower()
            engine = "openpyxl" if ext in ("xlsx", "xlsm") else "xlrd"

            xls = pd.ExcelFile(BytesIO(raw_bytes), engine=engine)

            file_got_data = False
            for sheet in xls.sheet_names:
                df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
                result, header_idx = extract_sheet(df_raw, file_name=file_name)
                label = f"{file_name} › {sheet}"
                if result is not None and not result.empty:
                    frames.append(result)
                    logs.append(
                        f"✅ {label}: tiêu đề bảng ở dòng {header_idx + 1}, "
                        f"{len(result)} học sinh"
                    )
                    file_got_data = True
                else:
                    logs.append(f"⚠️ {label}: không tìm thấy cột Họ tên, bỏ qua")

            if not file_got_data:
                logs.append(f"❌ {file_name}: không có sheet nào hợp lệ")

        except Exception as e:
            logs.append(f"❌ {file_name}: lỗi – {e}")

    if not frames:
        return pd.DataFrame(), logs

    # Gộp giữ nguyên toàn bộ dòng — KHÔNG lọc/xoá trùng
    merged = pd.concat(frames, ignore_index=True)

    # Sắp xếp cột: cột ưu tiên trước, còn lại giữ nguyên thứ tự
    existing_priority = [c for c in PRIORITY_COLS if c in merged.columns]
    other_cols = [c for c in merged.columns if c not in PRIORITY_COLS]
    merged = merged[existing_priority + other_cols]

    return merged, logs


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Xuất toàn bộ dữ liệu vào 1 sheet duy nhất tên 'Tổng hợp'.
    Cột 'Lớp' là tên file đầu vào (đã gán lúc extract)."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tổng hợp")
    return buf.getvalue()
