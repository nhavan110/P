# ── IMPORTS ──────────────────────────────────────────────────────────────────
import json
from copy import copy
from datetime import datetime, date
from pathlib import Path
from zoneinfo import ZoneInfo

VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")

import numpy as np
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter
from vnstock import Vnstock

# ── CẤU HÌNH ─────────────────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).resolve().parent
EXCEL_FILE   = BASE_DIR / "data" / "Portfolio.xlsx"
OUTPUT_FILE  = BASE_DIR / "data" / "MyPortfolio.xlsx"
JSON_FILE    = BASE_DIR / "data" / "data.json"   # trong thư mục data/, khớp với fetch('data/data.json') trong index.html và được workflow "git add data/" commit
MANUAL_FILE  = BASE_DIR / "manual_entries.csv"
SHEET_NAME   = "Sheet1"
VNI_START    = "2022-06-01"
API_TEMPLATE = "https://api.simplize.vn/api/historical/quote/prices/{}?page=0&size=600"

RISK_FREE_RATE = 0.04
TRADING_DAYS   = 252


# ── 1. ĐỌC DỮ LIỆU GỐC ───────────────────────────────────────────────────────
df_pr = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
stock_codes = [c for c in df_pr.columns if len(str(c)) == 3]


# ── 2. LẤY GIÁ CỔ PHIẾU TỪ API ───────────────────────────────────────────────
def fetch_stock_prices(codes: list) -> pd.DataFrame:
    """Lấy giá đóng cửa của danh sách mã CK, trả về DataFrame indexed by date."""
    frames = []
    for code in codes:
        resp = requests.get(API_TEMPLATE.format(code), timeout=10)
        resp.raise_for_status()
        df = pd.DataFrame(resp.json()["data"])
        df["date"] = pd.to_datetime(df["date"], unit="s")
        frames.append(df.set_index("date")["priceClose"].rename(code))
    return pd.concat(frames, axis=1).reset_index()   # outer join tự động


df_prices = fetch_stock_prices(stock_codes)


# ── 3. LẤY VNINDEX ────────────────────────────────────────────────────────────
today_str = date.today().strftime("%Y-%m-%d")
stock = Vnstock().stock(symbol="VNINDEX", source="VCI")
df_vni = (
    stock.quote.history(start=VNI_START, end=today_str)
    .rename(columns={"time": "date", "close": "VNINDEX"})
    .assign(date=lambda x: pd.to_datetime(x["date"]))
    .iloc[::-1]
    [["date", "VNINDEX"]]
)

df_prices["date"] = pd.to_datetime(df_prices["date"])
df_all = df_prices.merge(df_vni, on="date", how="left")


# ── 4. NỐI VỚI DỮ LIỆU CŨ ───────────────────────────────────────────────────
df_pr["date"] = pd.to_datetime(df_pr["date"], format="%d/%m/%Y")
last_update   = df_pr["date"].iloc[0]

df_new  = df_all[df_all["date"] > last_update]
df_pr   = (
    pd.concat([df_new, df_pr], ignore_index=True)
    .drop(columns=["Unnamed: 5", "Unnamed: 10"], errors="ignore")
    .fillna(0)
)


# ── 5. CẬP NHẬT W / D TỪ FILE (thay cho input() thủ công) ───────────────────
# Trước đây bước này hỏi trực tiếp trên terminal. Vì chạy tự động trên GitHub
# Actions không có ai ngồi gõ, ta đọc từ manual_entries.csv thay thế.
# Khi cần thêm W/D: mở file manual_entries.csv trên GitHub (web), thêm dòng
# "date,type,value" (vd: 12/08/2026,D,10000000), rồi để lần chạy tiếp theo tự
# áp dụng và tự xoá dòng đã xử lý.
df_pr["date_str"] = df_pr["date"].dt.strftime("%d/%m/%Y")


def apply_manual_entries(df: pd.DataFrame, manual_file: Path) -> pd.DataFrame:
    if not manual_file.exists():
        return df
    try:
        manual_df = pd.read_csv(manual_file)
    except pd.errors.EmptyDataError:
        return df
    if manual_df.empty:
        return df

    for _, row in manual_df.iterrows():
        date_str = str(row["date"]).strip()
        option   = str(row["type"]).strip().upper()
        if option not in ("W", "D"):
            print(f"⚠️  Bỏ qua dòng không hợp lệ (type='{option}')")
            continue
        try:
            value = float(row["value"])
        except (ValueError, TypeError):
            print(f"⚠️  Bỏ qua dòng không hợp lệ (value='{row['value']}')")
            continue

        mask = df["date_str"] == date_str
        if not mask.any():
            print(f"⚠️  Không tìm thấy ngày {date_str} trong dữ liệu, bỏ qua")
            continue
        df.loc[mask, option] = value
        print(f"✅ Đã cập nhật {option} ngày {date_str} = {value:,.0f}")

    # Xoá các dòng đã xử lý, chỉ giữ lại header
    pd.DataFrame(columns=["date", "type", "value"]).to_csv(manual_file, index=False)
    return df


df_pr = apply_manual_entries(df_pr, MANUAL_FILE)


# ── 6. TÍNH E1 (VECTORIZED) ───────────────────────────────────────────────────
def get_holdings(dates: pd.Series) -> pd.DataFrame:
    """
    Trả về DataFrame các cột b,c,d,e,f,g,h,margin theo từng ngày.
    ⚠️ MỖI LẦN GIAO DỊCH: thêm 1 dòng mới vào breakpoints bên dưới
    (sửa trực tiếp file này trên GitHub web, không cần chạy code ở máy).
    """
    breakpoints = [
        (datetime(2022,  9, 15), datetime(2023,  7,  6), dict(b=100)),
        (datetime(2023,  7,  7), datetime(2023, 10,  3), dict(b=100, c=100)),
        (datetime(2023, 10,  4), datetime(2023, 10, 20), dict(b=100, c=100, d=100)),
        (datetime(2023, 10, 23), datetime(2023, 10, 30), dict(b=100, c=100, d=100, e=100)),
        (datetime(2023, 10, 31), datetime(2024,  4,  1), dict(b=100, c=100, d=100, e=200)),
        (datetime(2024,  4,  2), datetime(2024,  5, 22), dict(b=100, d=100, e=200)),
        (datetime(2024,  5, 23), datetime(2024,  6,  6), dict(b=110, d=100, e=200)),
        (datetime(2024,  6,  7), datetime(2024,  6, 19), dict(b=110, d=100, e=200, f=100)),
        (datetime(2024,  6, 20), datetime(2024,  8,  5), dict(b=110, d=200, e=200, f=100)),
        (datetime(2024,  8,  6), datetime(2024,  8, 11), dict(b=110, d=200, e=100, f=100)),
        (datetime(2024,  8, 12), datetime(2024,  9,  4), dict(b=110, d=300, e=100, f=100)),
        (datetime(2024,  9,  5), datetime(2024,  9, 18), dict(b=210, d=300, e=100, f=100)),
        (datetime(2024,  9, 19), datetime(2025,  1,  6), dict(b=210, d=400, e=100, f=100)),
        (datetime(2025,  1,  7), datetime(2025,  3, 23), dict(b=210, d=400, e=100, f=115)),
        (datetime(2025,  3, 24), datetime(2025,  4,  3), dict(b=210, d=400, e=100)),
        (datetime(2025,  4,  4), datetime(2025,  4,  8), dict(b=310, d=400, e=100)),
        (datetime(2025,  4,  9), datetime(2025,  6, 12), dict(b=410, d=500, e=200)),
        (datetime(2025,  6, 13), datetime(2025,  7,  6), dict(b=492, d=500, e=200)),
        (datetime(2025,  7,  7), datetime(2025,  7, 10), dict(b=592, d=500, e=200, margin=2345000)),
        (datetime(2025,  7, 11), datetime(2025,  7, 17), dict(b=392, d=500, e=200)),
        (datetime(2025,  7, 18), datetime(2025,  8, 11), dict(b=392, d=100, e=200)),
        (datetime(2025,  8, 12), datetime(2025,  9, 18), dict(b=192, d=100, e=200)),
        (datetime(2025,  9, 19), datetime(2025,  9, 25), dict(b=192, d=100, e=200, g=100)),
        (datetime(2025,  9, 26), datetime(2025,  9, 28), dict(b=300, d=100, e=400, g=100, margin=8391768)),
        (datetime(2025,  9, 29), datetime(2025,  9, 29), dict(b=200, d=100, e=400, g=100, margin=8391768)),
        (datetime(2025,  9, 30), datetime(2025, 10, 13), dict(b=200, d=100, e=200, g=100, margin=8291768)),
        (datetime(2025, 10, 14), datetime(2025, 12, 14), dict(b=400, d=100, e=200, g=100, margin=14081631)),
        (datetime(2025, 12, 15), datetime(2026,  1, 21), dict(b=400, d=100, e=200, g=100, h=100, margin=23446444)),
        (datetime(2026,  1, 22), datetime(2026,  1, 25), dict(b=400, d=100, e=0,   g=100, h=100, margin=23446783)),
        (datetime(2026,  1, 26), datetime(2026,  2, 23), dict(b=400, d=100, g=100, h=100, margin=15000000)),
        (datetime(2026,  2, 24), datetime(2026,  3,  1), dict(b=400, d=100, g=100, h=300, margin=22000000)),
        (datetime(2026,  3,  2), datetime(2026,  3, 26), dict(b=400, d=100, g=0,   h=300, margin=16000000)),
        (datetime(2026,  3, 27), datetime(2026,  5, 10), dict(b=400, d=100, h=300, margin=8000000)),
        (datetime(2026,  5, 11), datetime(2026,  5, 24), dict(b=400, d=100, h=400, margin=15007665)),
        (datetime(2026,  5, 25), datetime(2026,  6,  2), dict(b=440, d=100, h=400, margin=15007665)),
        (datetime(2026,  6,  3), datetime(2026,  6,  9), dict(b=440, c=100, d=100, h=400, margin=17810000)),
        (datetime(2026,  6, 10), datetime(2026,  7,  5), dict(b=440, c=100, d=100, h=400, margin=17363606)),
        (datetime(2026,  7,  6), datetime(2026,  7,  7), dict(b=440, c=300, d=100, h=400, margin=22000000)),
        (datetime(2026,  7,  8), datetime(2026,  7,  8), dict(b=440, c=1100, d=100, h=400, margin=36000000)),
        (datetime(2026,  7,  9), datetime(2026,  7, 22), dict(b=440, c=1600, d=100, h=400, margin=43000000)),
        (datetime(2026,  7, 23), datetime(2026,  7, 26), dict(b=440, c=1600, d=100, h=400, margin=33000000)),
        (datetime(2026,  7, 27), datetime(2099,  7,  8), dict(b=440, c=1600, d=100, h=400, margin=28000000)),
    ]
    cols = ["b", "c", "d", "e", "f", "g", "h", "margin"]
    result = pd.DataFrame(0, index=dates.index, columns=cols)
    for start, end, vals in breakpoints:
        mask = (dates >= start) & (dates <= end)
        for col, val in vals.items():
            result.loc[mask, col] = val
    return result


dates_dt = pd.to_datetime(df_pr["date"])
holdings = get_holdings(dates_dt)

df_pr["E1"] = (
      df_pr["HPG"] * holdings["b"]
    - holdings["margin"]
    + df_pr["PNJ"] * holdings["c"]
    + df_pr["TCB"] * holdings["d"]
    + df_pr["MWG"] * holdings["e"]
    + df_pr["MBB"] * holdings["f"]
    + df_pr["FRT"] * holdings["g"]
    + df_pr["FPT"] * holdings["h"]
)


# ── 7. TÍNH CÁC CHỈ SỐ HIỆU SUẤT ─────────────────────────────────────────────
df_pr["E0"] = df_pr["E1"].shift(-1)
df_pr.loc[df_pr.index[-1], "E0"] = 0

df_pr["DR"]        = ((df_pr["E1"] + df_pr["W"]) - (df_pr["E0"] + df_pr["D"])) / (df_pr["E0"] + df_pr["D"])
df_pr["DR(VNI)"]   = df_pr["VNINDEX"].div(df_pr["VNINDEX"].shift(-1)).sub(1).fillna(0)
df_pr["DR+1"]      = df_pr["DR"] + 1
df_pr["DR+1(VNI)"] = df_pr["DR(VNI)"] + 1

df_pr["CR+1"]      = df_pr["DR+1"][::-1].cumprod()[::-1]
df_pr["CR+1(VNI)"] = df_pr["DR+1(VNI)"][::-1].cumprod()[::-1]
df_pr["CR"]        = df_pr["CR+1"] - 1
df_pr["CR(VNI)"]   = df_pr["CR+1(VNI)"] - 1


# ── 7b. TÍNH SHARPE RATIO & MAX DRAWDOWN DẠNG SỐ (mới, cho JSON dashboard) ──
df_pr["Sharpe"]      = np.nan
df_pr["MaxDrawdown"] = np.nan

# Cần tính năm (Year) trước để group
df_pr["Year_tmp"] = pd.to_datetime(df_pr["date"], format="%d/%m/%Y").dt.year
df_pr["YR"] = df_pr.groupby("Year_tmp")["DR+1"].transform(lambda s: s.cumprod().iloc[-1] - 1)

for _, idx in df_pr.groupby("YR").groups.items():
    idx = sorted(idx)
    dr_slice = df_pr.loc[idx, "DR"]
    std = dr_slice.std(ddof=0)
    if std:
        sharpe = (df_pr.loc[idx[0], "YR"] - RISK_FREE_RATE) / (std * np.sqrt(TRADING_DAYS))
        df_pr.loc[idx, "Sharpe"] = sharpe

    cr1_slice = pd.to_numeric(df_pr.loc[idx, "CR+1"], errors="coerce")[::-1]
    peaks = cr1_slice.cummax()
    max_dd = ((cr1_slice - peaks) / peaks).min()
    df_pr.loc[idx, "MaxDrawdown"] = max_dd


# ── 8. GHI EXCEL (Sheet1) ────────────────────────────────────────────────────
df_pr["date"] = pd.to_datetime(df_pr["date"]).dt.strftime("%d/%m/%Y")
df_pr = df_pr.drop(columns=["date_str"], errors="ignore")

cols_to_drop = ["DR+1", "DR+1(VNI)", "CR+1", "CR+1(VNI)", "CR(VNI)",
                "MonthYear", "Cumulative_DR_M", "Year", "Year_tmp",
                "Cumulative_DR_Y", "Cumulative_DR_Y(VNI)"]
df_save = df_pr.drop(columns=[c for c in cols_to_drop if c in df_pr.columns])
EXCEL_FILE.parent.mkdir(parents=True, exist_ok=True)
df_save.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)


# ── 9. THỐNG KÊ THÁNG / NĂM ──────────────────────────────────────────────────
df_pr["date"] = pd.to_datetime(df_pr["date"], format="%d/%m/%Y")

df_pr["MonthYear"] = df_pr["date"].dt.to_period("M")
df_pr["MR"] = df_pr.groupby("MonthYear")["DR+1"].transform(lambda s: s.cumprod().iloc[-1] - 1)

df_pr["Year"] = df_pr["date"].dt.year
df_pr["YR(VNI)"] = df_pr.groupby("Year")["DR+1(VNI)"].transform(lambda s: s.cumprod().iloc[-1] - 1)

df_pr["date"] = df_pr["date"].dt.strftime("%d/%m/%Y")


# ── 10. XUẤT MyPortfolio.xlsx ─────────────────────────────────────────────────
# Thứ tự cột cố định theo yêu cầu — "Sharpe ratio" và "Max drawdown" (dạng công
# thức Excel) được chèn thêm ở bước 11a ngay sau khi ghi file này, nên ở đây
# chỉ cần các cột số liệu gốc, không lấy 2 cột Sharpe/MaxDrawdown dạng số
# (chỉ dùng riêng cho JSON dashboard).
EXPORT_COLUMNS = [
    "date", "HPG", "TCB", "FPT", "PNJ", "FRT", "MWG", "MBB", "VNINDEX",
    "E1", "W", "E0", "D", "DR", "DR(VNI)", "CR", "MR", "YR", "YR(VNI)",
]
df_export = df_pr[[c for c in EXPORT_COLUMNS if c in df_pr.columns]].copy()
df_export.to_excel(OUTPUT_FILE, index=False)


# ── 11. ĐỊNH DẠNG OPENPYXL ───────────────────────────────────────────────────
wb = load_workbook(OUTPUT_FILE)
ws = wb["Sheet1"]


def header_map(sheet) -> dict:
    return {str(c.value).strip(): c.column for c in sheet[1] if c.value}


hm = header_map(ws)


def merge_same_values(sheet, col_name: str, df_col: pd.Series, hm: dict):
    col_idx = hm.get(col_name)
    if col_idx is None:
        return
    letter = get_column_letter(col_idx)
    for val, idxs in df_col.groupby(df_col).groups.items():
        idxs = sorted(idxs)
        r1, r2 = idxs[0] + 2, idxs[-1] + 2
        sheet.merge_cells(f"{letter}{r1}:{letter}{r2}")


for col in ("MR", "YR", "YR(VNI)"):
    merge_same_values(ws, col, df_export[col] if col in df_export.columns else df_pr[col], hm)

yr_vni_idx = hm.get("YR(VNI)")
dr_idx     = hm.get("DR")
if yr_vni_idx and dr_idx:
    sharpe_idx    = yr_vni_idx + 1
    sharpe_letter = get_column_letter(sharpe_idx)
    dr_letter     = get_column_letter(dr_idx)
    ws[f"{sharpe_letter}1"] = "Sharpe ratio"

    df_yr = df_export["YR"] if "YR" in df_export.columns else df_pr["YR"]
    dr_col = df_export["DR"] if "DR" in df_export.columns else df_pr["DR"]
    for val, idxs in df_yr.groupby(df_yr).groups.items():
        idxs = sorted(idxs)
        r1, r2 = idxs[0] + 2, idxs[-1] + 2
        ws.merge_cells(f"{sharpe_letter}{r1}:{sharpe_letter}{r2}")
        ws[f"{sharpe_letter}{r1}"] = (
            f"=({val} - {RISK_FREE_RATE}) / "
            f"(STDEVP({dr_letter}{r1}:{dr_letter}{r2}) * SQRT({TRADING_DAYS}))"
        )

if yr_vni_idx:
    dd_idx    = yr_vni_idx + 2
    dd_letter = get_column_letter(dd_idx)
    ws[f"{dd_letter}1"] = "Max drawdown"

    df_cr1 = (df_export["CR+1"] if "CR+1" in df_export.columns else df_pr["CR+1"]).copy()
    df_yr2 = df_export["YR"] if "YR" in df_export.columns else df_pr["YR"]
    for val, idxs in df_yr2.groupby(df_yr2).groups.items():
        idxs = sorted(idxs)
        r1, r2 = idxs[0] + 2, idxs[-1] + 2
        ws.merge_cells(f"{dd_letter}{r1}:{dd_letter}{r2}")
        cr1_slice = pd.to_numeric(df_cr1.iloc[idxs[0]: idxs[-1] + 2], errors="coerce")[::-1]
        peaks     = cr1_slice.cummax()
        max_dd    = ((cr1_slice - peaks) / peaks).min()
        ws[f"{dd_letter}{r1}"] = max_dd


hm = header_map(ws)

date_style = NamedStyle(name="date_fmt", number_format="DD-MM-YYYY")
int_acc    = NamedStyle(name="int_acc",  number_format='_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)')
pct_style  = NamedStyle(name="pct_2",    number_format="0.00%")


def apply_style_to_cols(sheet, col_names, style, hm):
    for name in col_names:
        ci = hm.get(name)
        if ci:
            for row in sheet.iter_rows(min_row=2, min_col=ci, max_col=ci):
                for cell in row:
                    cell.style = style


date_col = hm.get("date")
if date_col:
    for row in ws.iter_rows(min_row=2, min_col=date_col, max_col=date_col):
        for cell in row:
            cell.style = date_style

apply_style_to_cols(ws, ["HPG", "TCB", "MWG", "MBB", "PNJ", "FRT", "FPT", "VNINDEX", "E1", "W", "E0", "D"], int_acc, hm)
apply_style_to_cols(ws, ["DR", "DR(VNI)", "CR", "MR", "YR", "YR(VNI)", "Max drawdown"], pct_style, hm)

sharpe_ci = hm.get("Sharpe ratio") or (yr_vni_idx + 1 if yr_vni_idx else None)
if sharpe_ci:
    for row in ws.iter_rows(min_row=1, min_col=sharpe_ci, max_col=sharpe_ci):
        for cell in row:
            cell.number_format = "0.00"

thin = Border(left=Side(style="thin"), right=Side(style="thin"),
              top=Side(style="thin"),  bottom=Side(style="thin"))
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        if cell.value is not None:
            cell.border = thin

COLOR_MAP = {
    "date": "C0FFC0", "HPG": "C0FFC0", "TCB": "C0FFC0", "FPT": "C0FFC0", "PNJ": "C0FFC0",
    "FRT": "CCFFFF", "MWG": "CCFFFF", "MBB": "CCFFFF",
    "VNINDEX": "FFA500",
    "E1": "4F81BD", "W": "4F81BD", "E0": "4F81BD", "D": "4F81BD",
    "DR": "8E7CC3", "DR(VNI)": "8E7CC3", "CR": "8E7CC3", "MR": "8E7CC3",
    "YR": "8E7CC3", "YR(VNI)": "8E7CC3", "Sharpe ratio": "8E7CC3", "Max drawdown": "8E7CC3",
}

hm = header_map(ws)
for col_name, hex_color in COLOR_MAP.items():
    ci = hm.get(col_name)
    if ci:
        fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
        for row in ws.iter_rows(min_row=1, min_col=ci, max_col=ci):
            for cell in row:
                cell.fill = fill

src = ws["A1"]
for col in range(2, ws.max_column + 1):
    tgt = ws.cell(row=1, column=col)
    tgt.font      = copy(src.font)
    tgt.alignment = copy(src.alignment)

COLUMN_WIDTHS = {
    "date": 10.6, "HPG": 8.6, "TCB": 8.6, "MWG": 8.6, "FRT": 8.6,
    "FPT": 8.6, "MBB": 8.6, "PNJ": 8.6,
    "VNINDEX": 8.8, "E1": 11.0, "W": 11.0, "E0": 11.0, "D": 11.0,
    "DR": 8.1, "DR(VNI)": 8.1, "CR": 8.1, "MR": 8.1,
    "YR": 8.1, "YR(VNI)": 8.1, "Sharpe ratio": 11.2, "Max drawdown": 14.2,
}
hm = header_map(ws)
for name, width in COLUMN_WIDTHS.items():
    ci = hm.get(name)
    if ci:
        ws.column_dimensions[get_column_letter(ci)].width = width

wb.save(OUTPUT_FILE)
print(f"✅ Đã lưu: {OUTPUT_FILE}")


# ── 12. XUẤT JSON CHO WEB DASHBOARD ──────────────────────────────────────────
def safe_num(x):
    """Trả về float hợp lệ, hoặc None nếu là NaN/Infinity (JSON chuẩn không
    chấp nhận NaN/Infinity — trình duyệt sẽ báo lỗi parse nếu để lọt vào)."""
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    if not np.isfinite(v):
        return None
    return v


def export_json(df: pd.DataFrame, path: Path) -> None:
    records = df.copy()
    records["date"] = pd.to_datetime(records["date"], format="%d/%m/%Y")
    records = records.sort_values("date")
    records["date"] = records["date"].dt.strftime("%Y-%m-%d")

    keep_cols = [c for c in [
        "date", "HPG", "TCB", "FPT", "PNJ", "FRT", "MWG", "MBB", "VNINDEX",
        "E1", "W", "E0", "D", "DR", "DR(VNI)", "CR", "CR(VNI)",
        "MR", "YR", "YR(VNI)", "Sharpe", "MaxDrawdown",
    ] if c in records.columns]

    history = []
    for _, row in records[keep_cols].iterrows():
        rec = {"date": row["date"]}
        for c in keep_cols:
            if c == "date":
                continue
            rec[c] = safe_num(row[c])
        history.append(rec)

    latest = records.iloc[-1]
    summary = {
        "last_updated": datetime.now(VN_TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "portfolio_value": safe_num(latest.get("E1")),
        "cumulative_return": safe_num(latest.get("CR")),
        "cumulative_return_vni": safe_num(latest.get("CR(VNI)")),
        "sharpe_ratio": safe_num(latest.get("Sharpe")),
        "max_drawdown": safe_num(latest.get("MaxDrawdown")),
    }

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"summary": summary, "history": history}, f,
                   ensure_ascii=False, indent=2, allow_nan=False)
    print(f"✅ Đã xuất: {path}")


export_json(df_pr, JSON_FILE)
