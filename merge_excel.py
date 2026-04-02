import pandas as pd
import re
from io import BytesIO

# ─── Ánh xạ tên cột chuẩn ───────────────────────────────────────────────────
COLUMN_ALIASES = {
    "lop": ["lớp", "lop", "lớp học", "lop hoc", "class", "khối lớp", "khoi lop"],
    "ho_ten": [
        "họ tên", "ho ten", "họ và tên", "ho va ten", "tên", "ten",
        "full name", "name", "họ tên học sinh", "ho ten hoc sinh",
        "tên học sinh", "ten hoc sinh",
    ],
    "ngay_sinh": [
        "ngày sinh", "ngay sinh", "ngày tháng năm sinh", "ngay thang nam sinh",
        "dob", "date of birth", "năm sinh", "nam sinh",
    ],
}

STANDARD_NAMES = {
    "lop": "Lớp",
    "ho_ten": "Họ và tên",
    "ngay_sinh": "Ngày sinh",
}

REQUIRED_COLS = {"lop", "ho_ten", "ngay_sinh"}


def _normalize(text: str) -> str:
    """Chuẩn hoá chuỗi: lowercase, bỏ dấu cơ bản, trim."""
    text = str(text).strip().lower()
    # bỏ dấu tiếng Việt đơn giản bằng regex (không cần unidecode)
    replacements = [
        (r"[àáâãäåạảấầẩẫậắằẳẵặ]", "a"),
        (r"[èéêëẹẻẽếềểễệ]", "e"),
        (r"[ìíîïịỉĩ]", "i"),
        (r"[òóôõöọỏốồổỗộớờởỡợ]", "o"),
        (r"[ùúûüụủũứừửữự]", "u"),
        (r"[ỳýỹỵỷ]", "y"),
        (r"[đ]", "d"),
    ]
    for pat, rep in replacements:
        text = re.sub(pat, rep, text)
    return text


def _detect_header_row(df_raw: pd.DataFrame) -> int | None:
    """Tìm dòng header (dòng chứa ít nhất 2 trong 3 cột cần thiết)."""
    all_aliases = {k: [_normalize(a) for a in v] for k, v in COLUMN_ALIASES.items()}

    for i, row in df_raw.iterrows():
        row_vals = [_normalize(str(c)) for c in row.values]
        found = set()
        for std_key, aliases in all_aliases.items():
            for val in row_vals:
                if val in aliases:
                    found.add(std_key)
                    break
        if len(found) >= 2:
            return i
    return None


def _map_columns(header_row: pd.Series) -> dict:
    """Trả về {std_key: col_index} từ dòng header."""
    all_aliases = {k: [_normalize(a) for a in v] for k, v in COLUMN_ALIASES.items()}
    mapping = {}
    for col_idx, cell in enumerate(header_row):
        cell_norm = _normalize(str(cell))
        for std_key, aliases in all_aliases.items():
            if cell_norm in aliases and std_key not in mapping:
                mapping[std_key] = col_idx
    return mapping


def extract_sheet(df_raw: pd.DataFrame, source_label: str) -> pd.DataFrame | None:
    """Trích xuất dữ liệu từ 1 sheet thô."""
    header_idx = _detect_header_row(df_raw)
    if header_idx is None:
        return None

    header_row = df_raw.iloc[header_idx]
    col_map = _map_columns(header_row)

    # Phải có đủ 3 cột mới lấy
    if not REQUIRED_COLS.issubset(col_map.keys()):
        return None

    data_rows = df_raw.iloc[header_idx + 1 :].reset_index(drop=True)

    records = []
    for _, row in data_rows.iterrows():
        ho_ten_val = str(row.iloc[col_map["ho_ten"]]).strip()
        # Bỏ dòng trống / dòng tổng hợp
        if not ho_ten_val or ho_ten_val.lower() in ("nan", "", "none"):
            continue

        record = {
            STANDARD_NAMES["lop"]: row.iloc[col_map["lop"]],
            STANDARD_NAMES["ho_ten"]: ho_ten_val,
            STANDARD_NAMES["ngay_sinh"]: row.iloc[col_map["ngay_sinh"]],
            "Nguồn": source_label,
        }
        records.append(record)

    return pd.DataFrame(records) if records else None


def merge_excel_files(uploaded_files: list) -> tuple[pd.DataFrame, list[str]]:
    """
    Nhận list các file object (có .name và .read()),
    trả về (DataFrame tổng hợp, danh sách log).
    """
    logs = []
    frames = []

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        try:
            raw_bytes = uploaded_file.read()
            ext = file_name.rsplit(".", 1)[-1].lower()

            # Đọc tất cả sheet
            xls = pd.ExcelFile(BytesIO(raw_bytes), engine="openpyxl" if ext in ("xlsx", "xlsm") else "xlrd")
            sheet_names = xls.sheet_names

            file_got_data = False
            for sheet in sheet_names:
                df_raw = pd.read_excel(BytesIO(raw_bytes), sheet_name=sheet, header=None,
                                       engine="openpyxl" if ext in ("xlsx", "xlsm") else "xlrd")
                label = f"{file_name} | {sheet}"
                result = extract_sheet(df_raw, label)
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

    if frames:
        merged = pd.concat(frames, ignore_index=True)
        merged = merged.drop_duplicates(subset=[STANDARD_NAMES["ho_ten"], STANDARD_NAMES["ngay_sinh"]])
        return merged, logs
    else:
        return pd.DataFrame(), logs


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tổng hợp")
    return buf.getvalue()
