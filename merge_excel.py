import pandas as pd
import re
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

PRIORITY_COLS = ["Họ và tên", "Lớp", "Giới tính", "Ngày sinh", "Lớp_gộp"]
REQUIRED_COLS = {"lop", "ho_ten", "ngay_sinh"}
HEADER_SCAN_ROWS = 10


def _normalize(text: str) -> str:
    text = str(text).strip().lower()
    replacements = [
        (r"[àáâãäåạảấầẩẫậắằẳẵặ]", "a"),
        (r"[èéêëẹẻẽếềểễệ]",       "e"),
        (r"[ìíîïịỉĩ]",             "i"),
        (r"[òóôõöọỏốồổỗộớờởỡợ]",  "o"),
        (r"[ùúûüụủũứừửữự]",        "u"),
        (r"[ỳýỹỵỷ]",               "y"),
        (r"[đ]",                   "d"),
    ]
    for pat, rep in replacements:
        text = re.sub(pat, rep, text)
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _detect_header_row(df_raw: pd.DataFrame) -> int | None:
    all_aliases = {k: [_normalize(a) for a in v] for k, v in COLUMN_ALIASES.items()}
    scan_limit = min(HEADER_SCAN_ROWS, len(df_raw))
    for i in range(scan_limit):
        row = df_raw.iloc[i]
        row_vals = [_normalize(str(c)) for c in row.values]
        found = set()
        for std_key, aliases in all_aliases.items():
            for val in row_vals:
                if val in aliases:
                    found.add(std_key)
                    break
        if len(found & REQUIRED_COLS) >= 2:
            return i
    return None


def _map_columns(header_row: pd.Series) -> dict:
    all_aliases = {k: [_normalize(a) for a in v] for k, v in COLUMN_ALIASES.items()}
    mapping = {}
    for col_idx, cell in enumerate(header_row):
        cell_norm = _normalize(str(cell))
        for std_key, aliases in all_aliases.items():
            if cell_norm in aliases and std_key not in mapping:
                mapping[std_key] = col_idx
    return mapping


def _format_date(val) -> str:
    if pd.isnull(val) if not isinstance(val, str) else False:
        return ""
    s = str(val).strip()
    if s.lower() in ("", "nan", "none", "nat"):
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime("%d/%m/%Y")
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return pd.to_datetime(s, format=fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return s


def extract_sheet(df_raw: pd.DataFrame, file_name: str) -> pd.DataFrame | None:
    header_idx = _detect_header_row(df_raw)
    if header_idx is None:
        return None

    header_row = df_raw.iloc[header_idx]
    col_map = _map_columns(header_row)

    if not REQUIRED_COLS.issubset(col_map.keys()):
        return None

    original_col_names = list(header_row.values)
    data_rows = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
    mapped_indices = set(col_map.values())

    # Tên file bỏ phần đuôi để dùng làm Lớp_gộp
    lop_gop_val = file_name.rsplit(".", 1)[0]

    records = []
    for _, row in data_rows.iterrows():
        ho_ten_val = str(row.iloc[col_map["ho_ten"]]).strip()
        if not ho_ten_val or ho_ten_val.lower() in ("nan", "", "none"):
            continue

        record = {
            STANDARD_NAMES["ho_ten"]:    ho_ten_val,
            STANDARD_NAMES["lop"]:       str(row.iloc[col_map["lop"]]).strip(),
            STANDARD_NAMES["gioi_tinh"]: (
                str(row.iloc[col_map["gioi_tinh"]]).strip()
                if "gioi_tinh" in col_map else ""
            ),
            STANDARD_NAMES["ngay_sinh"]: _format_date(row.iloc[col_map["ngay_sinh"]]),
            "Lớp_gộp":                   lop_gop_val,
        }

        # Giữ các cột gốc khác
        for ci, orig_name in enumerate(original_col_names):
            if ci in mapped_indices:
                continue
            col_label = str(orig_name).strip()
            if col_label.lower() in ("nan", "", "none"):
                col_label = f"Cột_{ci}"
            if col_label in PRIORITY_COLS:
                col_label = f"{col_label}_gốc"
            record[col_label] = row.iloc[ci]

        records.append(record)

    return pd.DataFrame(records) if records else None


def merge_excel_files(uploaded_files: list) -> tuple[pd.DataFrame, list[str]]:
    logs = []
    frames = []

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        try:
            raw_bytes = uploaded_file.read()
            ext = file_name.rsplit(".", 1)[-1].lower()
            engine = "openpyxl" if ext in ("xlsx", "xlsm") else "xlrd"

            xls = pd.ExcelFile(BytesIO(raw_bytes), engine=engine)
            sheet_names = xls.sheet_names

            file_got_data = False
            for sheet in sheet_names:
                df_raw = pd.read_excel(
                    BytesIO(raw_bytes), sheet_name=sheet,
                    header=None, engine=engine,
                )
                result = extract_sheet(df_raw, file_name=file_name)
                label = f"{file_name} › {sheet}"
                if result is not None and not result.empty:
                    frames.append(result)
                    logs.append(f"✅ {label}: {len(result)} học sinh")
                    file_got_data = True
                else:
                    logs.append(f"⚠️ {label}: không tìm thấy cột phù hợp, bỏ qua")

            if not file_got_data:
                logs.append(f"❌ {file_name}: không có sheet nào hợp lệ")

        except Exception as e:
            logs.append(f"❌ {file_name}: lỗi – {e}")

    if not frames:
        return pd.DataFrame(), logs

    merged = pd.concat(frames, ignore_index=True)
    merged = merged.drop_duplicates(
        subset=[STANDARD_NAMES["ho_ten"], STANDARD_NAMES["ngay_sinh"]]
    )

    # Sắp xếp cột: 4 ưu tiên trước, còn lại giữ nguyên
    existing_priority = [c for c in PRIORITY_COLS if c in merged.columns]
    other_cols = [c for c in merged.columns if c not in PRIORITY_COLS]
    merged = merged[existing_priority + other_cols]

    return merged, logs


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Xuất toàn bộ dữ liệu vào 1 sheet duy nhất tên 'Tổng hợp'.
    Cột 'Lớp' chứa tên sheet gốc của file input (đã gán lúc extract)."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tổng hợp")
    return buf.getvalue()
