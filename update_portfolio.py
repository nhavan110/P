# ── IMPORTS ──────────────────────────────────────────────────────────────────
import json
import sys
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
from scipy.optimize import brentq, newton
from vnstock import Vnstock

# ── WORKAROUND: bug trong vnstock.core.utils.env.get_hosting_service() ──────
# Hàm này không có nhánh else, nên khi chạy trên môi trường không phải
# Colab/Codespace/Replit/Kaggle/HF Spaces (vd: GitHub Actions runner), biến
# `hosting_service` không được gán -> UnboundLocalError. Lỗi này bị tenacity
# retry rồi che thành `tenacity.RetryError` khi gọi stock.quote.history().
# Patch lại hàm để trả về giá trị mặc định an toàn thay vì crash.
import vnstock.core.utils.env as _vnstock_env

_original_get_hosting_service = _vnstock_env.get_hosting_service


def _safe_get_hosting_service():
    try:
        result = _original_get_hosting_service()
    except UnboundLocalError:
        result = None
    return result or "Local or Unknown"


_vnstock_env.get_hosting_service = _safe_get_hosting_service

# ── CẤU HÌNH ─────────────────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).resolve().parent
PRICE_HISTORY_FILE     = BASE_DIR / "data" / "price_history.csv"      # giá thô (input/cache), KHÔNG chứa W/D hay kết quả
DNSE_DIR                = BASE_DIR / "data" / "dnse"                  # DNSE (chạy local) đẩy lên đây, xem DNSE/publish_to_github.py
HOLDINGS_FILE           = DNSE_DIR / "holdings_history.csv"           # SOURCE OF TRUTH: số lượng từng mã + D/W theo ngày
DNSE_POSITIONS_FILE     = DNSE_DIR / "positions.csv"                  # vị thế hiện tại của DNSE, chỉ để đối chiếu dòng mới nhất
PORTFOLIO_HISTORY_FILE  = BASE_DIR / "data" / "portfolio_history.csv"  # output: toàn bộ dữ liệu đã tính, ghi lại mỗi lần chạy
OUTPUT_FILE  = BASE_DIR / "data" / "MyPortfolio.xlsx"                  # output: bản Excel định dạng đẹp để xem/Drive
JSON_FILE    = BASE_DIR / "data" / "data.json"   # trong thư mục data/, khớp với fetch('data/data.json') trong index.html và được workflow "git add data/" commit
SHEET_NAME   = "Sheet1"
VNI_START    = "2022-06-01"
API_TEMPLATE = "https://api.simplize.vn/api/historical/quote/prices/{}?page=0&size=600"
NON_SYMBOL_COLS = {"date", "D", "W"}   # các cột còn lại của holdings_history.csv đều là mã CK

RISK_FREE_RATE = 0.04
TRADING_DAYS   = 252

# Map mã CP → ngành, dùng cho Allocation/Position Contribution (mục 4).
# Mã chưa có trong map sẽ hiện "Khác" (xem SECTOR_DEFAULT).
SECTOR_DEFAULT = "Khác"
SECTOR_MAP = {
    "HPG": "Thép", "TCB": "Ngân hàng", "MBB": "Ngân hàng",
    "FPT": "Công nghệ", "PNJ": "Bán lẻ", "FRT": "Bán lẻ", "MWG": "Bán lẻ",
}


# ── HÀM TIỆN ÍCH DÙNG CHUNG ───────────────────────────────────────────────────
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


# ── HÀM XIRR (mục 5) — dùng để tính lãi suất nội bộ trên chuỗi cashflow không
# đều theo thời gian. Đặt sớm vì không phụ thuộc df_pr, chỉ nhận list (date, amount).
def xirr(cashflows: list) -> float:
    """cashflows: list các (ngày, số tiền có dấu). Âm = tiền ra khỏi túi NĐT
    (D, hoặc E0 giả định mua vào đầu kỳ). Dương = tiền về túi NĐT
    (W, hoặc E1 giả định bán ra cuối kỳ)."""
    if len(cashflows) < 2:
        return None
    d0 = min(d for d, _ in cashflows)
    if all(d == d0 for d, _ in cashflows):   # mọi dòng tiền cùng 1 ngày -> XIRR không xác định
        return None

    def npv(rate):
        return sum(cf / (1 + rate) ** ((d - d0).days / 365) for d, cf in cashflows)

    try:
        return newton(npv, 0.1, maxiter=100)
    except (RuntimeError, OverflowError):
        try:
            return brentq(npv, -0.9999, 10)
        except ValueError:
            return None


# ── 0. ĐỌC holdings_history.csv (DNSE) — danh sách mã + số lượng + D/W ──────
def load_holdings_history(path: Path):
    """Đọc data/dnse/holdings_history.csv (ngày dd/mm/yyyy). Trả về (df tăng dần theo
    ngày, danh sách mã). Mã = mọi cột trừ date, D, W — danh sách động, không giới hạn."""
    if not path.exists():
        print(f"❌ Không tìm thấy {path}. DNSE (máy local) phải chạy publish_to_github.py trước.")
        sys.exit(1)
    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]
    df["date"] = pd.to_datetime(df["date"].astype(str).str.strip(), format="%d/%m/%Y")
    codes = [c for c in df.columns if c not in NON_SYMBOL_COLS]
    for c in codes + ["D", "W"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0) if c in df.columns else 0.0
    df = (df.drop_duplicates(subset="date", keep="last")
            .sort_values("date").reset_index(drop=True))
    return df, codes


holdings_hist, stock_codes = load_holdings_history(HOLDINGS_FILE)
if not stock_codes or holdings_hist[stock_codes].abs().sum().sum() == 0:
    print("ℹ️  holdings_history.csv chưa có mã nào đang/đã giữ — không có gì để tính. Bỏ qua lần chạy này.")
    sys.exit(0)
HISTORY_START = holdings_hist["date"].min()   # lịch sử + so sánh VNINDEX bắt đầu từ ngày này
print(f"📂 holdings_history.csv: {len(holdings_hist)} dòng, mã: {stock_codes}, bắt đầu {HISTORY_START:%d/%m/%Y}")


# ── 1. ĐỌC LỊCH SỬ GIÁ (price_history.csv — chỉ giá thô) ────────────────────
def load_price_history(path: Path) -> pd.DataFrame:
    cols = ["date", *stock_codes, "VNINDEX"]
    if not path.exists():
        return pd.DataFrame(columns=cols).astype({"date": "datetime64[ns]"})
    df = pd.read_csv(path)
    df["date"] = pd.to_datetime(df["date"], format="%d/%m/%Y")
    return df.sort_values("date").reset_index(drop=True)


price_hist = load_price_history(PRICE_HISTORY_FILE)


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


# ── 4. NỐI VỚI DỮ LIỆU CŨ, LƯU LẠI price_history.csv (chỉ giá thô) ──────────
last_update = price_hist["date"].max() if not price_hist.empty else (pd.Timestamp(VNI_START) - pd.Timedelta(days=1))

# Mã mới (chưa có cột trong price_history.csv): lấy giá lịch sử từ API để bù các ngày cũ.
_new_codes = [c for c in stock_codes if c not in price_hist.columns]
for _c in _new_codes:
    price_hist[_c] = price_hist["date"].map(df_prices.set_index("date")[_c])
if _new_codes:
    print(f"➕ Mã mới, đã bù giá lịch sử từ API: {_new_codes}")

df_new     = df_all[df_all["date"] > last_update]
price_hist = (
    pd.concat([price_hist, df_new], ignore_index=True)
    .drop_duplicates(subset="date", keep="last")
    .sort_values("date")
    .fillna(0)
    .reset_index(drop=True)
)

price_hist_out = price_hist.copy()
price_hist_out["date"] = price_hist_out["date"].dt.strftime("%d/%m/%Y")
PRICE_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
price_hist_out.to_csv(PRICE_HISTORY_FILE, index=False)
print(f"✅ Đã lưu: {PRICE_HISTORY_FILE} ({len(price_hist_out)} dòng)")

# df_pr = bảng làm việc trong bộ nhớ (mới nhất trước, khớp quy ước cũ) — không
# ghi đè price_history.csv, các cột tính toán bên dưới chỉ tồn tại trong biến
# này rồi xuất ra portfolio_history.csv / MyPortfolio.xlsx / data.json.
df_pr = price_hist.sort_values("date", ascending=False).reset_index(drop=True)


# ── 5. LỊCH SỬ BẮT ĐẦU TỪ NGÀY ĐẦU TIÊN CỦA holdings_history.csv ────────────
# price_history.csv vẫn giữ nguyên toàn bộ giá cũ, nhưng danh mục/VNINDEX/lợi nhuận/
# Sharpe/XIRR chỉ tính từ HISTORY_START trở đi.
df_pr = df_pr[df_pr["date"] >= HISTORY_START].reset_index(drop=True)
if df_pr.empty:
    print(f"ℹ️  Chưa có phiên giao dịch nào từ {HISTORY_START:%d/%m/%Y} — bỏ qua lần chạy này.")
    sys.exit(0)

df_pr["date_str"] = df_pr["date"].dt.strftime("%d/%m/%Y")
trading_days = np.sort(df_pr["date"].values)   # datetime64, tăng dần


def next_trading_day(d):
    """Phiên giao dịch đầu tiên >= ngày d (D/W rơi vào ngày nghỉ được dồn sang phiên kế tiếp)."""
    i = np.searchsorted(trading_days, np.datetime64(d), side="left")
    return pd.Timestamp(trading_days[i]) if i < len(trading_days) else None


# ── 5b. SỐ LƯỢNG TỪNG MÃ THEO NGÀY: lấy dòng gần nhất <= ngày đó (P tự điền các ngày giữa)
_tmp = df_pr[["date"]].copy()
_tmp["_i"] = df_pr.index
holdings = (
    pd.merge_asof(_tmp.sort_values("date"), holdings_hist[["date", *stock_codes]],
                  on="date", direction="backward")
    .set_index("_i").reindex(df_pr.index)[stock_codes].fillna(0.0)
)

# ── 5c. D / W TỪ holdings_history.csv ───────────────────────────────────────
# Mỗi dòng có D hoặc W được gán vào phiên giao dịch đầu tiên >= ngày dòng đó.
_dw = {}   # phiên -> [D, W]
_unmapped = []
for _, r in holdings_hist.iterrows():
    if not (r["D"] or r["W"]):
        continue
    day = next_trading_day(r["date"])
    if day is None:
        _unmapped.append(r["date"].strftime("%d/%m/%Y"))
        continue
    slot = _dw.setdefault(day, [0.0, 0.0])
    slot[0] += float(r["D"])
    slot[1] += float(r["W"])
if _unmapped:
    print(f"⚠️  D/W ở các ngày chưa có phiên giao dịch (chưa có giá): {_unmapped} — chưa được tính.")

# ── 6. TÍNH E1 (VECTORIZED): E1 = Σ giá × số lượng ──────────────────────────
df_pr["E1"] = (df_pr[stock_codes] * holdings[stock_codes]).sum(axis=1)

# Ngày đầu có sẵn cổ phiếu mà không có lệnh mua (mua trước DNSE_FROM_DATE):
# coi giá trị thị trường của số cổ phiếu đó là vốn nạp ban đầu (D) của ngày đầu.
_first_idx = df_pr["date"].idxmin()
_first_day = df_pr.loc[_first_idx, "date"]
_first_e1  = float(df_pr.loc[_first_idx, "E1"])
if _first_e1 > 0 and _dw.get(_first_day, [0.0, 0.0])[0] == 0:
    _dw.setdefault(_first_day, [0.0, 0.0])[0] = _first_e1
    print(f"ℹ️  Ngày đầu {_first_day:%d/%m/%Y} không có lệnh mua: coi giá trị cổ phiếu {_first_e1:,.0f} là vốn nạp ban đầu (D).")

# Bảng D/W theo (ngày, loại) — dùng chung cho cột D/W, XIRR và events
_cf_rows = []
for day, (d_val, w_val) in sorted(_dw.items()):
    if d_val:
        _cf_rows.append((day.strftime("%d/%m/%Y"), "D", d_val))
    if w_val:
        _cf_rows.append((day.strftime("%d/%m/%Y"), "W", w_val))
cashflows = pd.DataFrame(_cf_rows, columns=["date", "type", "value"])

if not cashflows.empty:
    wd = (
        cashflows.groupby(["date", "type"])["value"].sum()
        .unstack(fill_value=0.0)
        .reindex(columns=["W", "D"], fill_value=0.0)
    )
else:
    wd = pd.DataFrame(columns=["W", "D"])

df_pr = df_pr.merge(wd, left_on="date_str", right_index=True, how="left")
df_pr[["W", "D"]] = df_pr[["W", "D"]].fillna(0.0)

# Đối chiếu dòng mới nhất của holdings_history.csv với positions.csv của DNSE
def check_against_positions(path: Path, hold_hist: pd.DataFrame, codes: list) -> None:
    if not path.exists():
        print("ℹ️  Không có data/dnse/positions.csv để đối chiếu.")
        return
    pos = pd.read_csv(path, encoding="utf-8-sig")
    if pos.empty or "vị thế" not in pos.columns:
        actual = {}
    else:
        actual = dict(zip(pos["vị thế"], pd.to_numeric(pos["Khối lượng cộng dồn"], errors="coerce").fillna(0.0)))
    latest = hold_hist.iloc[-1]
    bad = {c: (float(latest[c]), float(actual.get(c, 0.0)))
           for c in codes if abs(float(latest[c]) - float(actual.get(c, 0.0))) > 1e-9}
    bad.update({c: (0.0, float(q)) for c, q in actual.items() if c not in codes and q})
    if bad:
        print(f"⚠️  holdings_history (dòng mới nhất) lệch positions.csv: {bad}  (mã: holdings, positions)")
    else:
        print("✅ holdings_history (dòng mới nhất) khớp positions.csv.")


check_against_positions(DNSE_POSITIONS_FILE, holdings_hist, stock_codes)


# ── 6b. SỰ KIỆN GIAO DỊCH GỘP THEO NGÀY CHO MARKER + TOOLTIP TRÊN CHART ──────
# Dựng từ holdings_history.csv: so số lượng mỗi dòng với dòng trước đó (dòng đầu
# so với 0). Mỗi phiên giao dịch có đổi số lượng -> 1 event:
#   - stocks / corp_stocks : {symbol: delta_qty}, dấu +/- theo tăng/giảm
#   - cash_in / cash_out   : D / W của phiên đó
# Quy tắc: đổi số lượng mà KHÔNG có tiền (D = W = 0) là cổ phiếu thưởng/chia tách
# (category "corporate_action", đưa vào corp_stocks); còn lại là "trade".
# Ngày chỉ có D/W mà số lượng không đổi (cổ tức, phí lưu ký) không tạo event.
def build_transaction_events(hold: pd.DataFrame, codes: list, dw_by_day: dict) -> list:
    by_day = {}
    prev = {c: 0.0 for c in codes}
    for _, row in hold.iterrows():
        cur = {c: float(row[c]) for c in codes}
        deltas = {c: cur[c] - prev[c] for c in codes if cur[c] != prev[c]}
        prev = cur
        if not deltas:
            continue
        day = next_trading_day(row["date"])
        if day is None:
            continue
        slot = by_day.setdefault(day, {})
        for c, q in deltas.items():
            slot[c] = slot.get(c, 0.0) + q

    events = []
    for day, slot in by_day.items():
        stocks = {c: q for c, q in slot.items() if q != 0}
        if not stocks:
            continue
        d_val, w_val = dw_by_day.get(day, [0.0, 0.0])
        category = "trade" if (d_val or w_val) else "corporate_action"
        event = {
            "date":        day.strftime("%Y-%m-%d"),
            "category":    category,
            "stocks":      {},
            "corp_stocks": {},
            "cash_in":     safe_num(d_val) if d_val else 0.0,
            "cash_out":    safe_num(w_val) if w_val else 0.0,
        }
        target_key = "corp_stocks" if category == "corporate_action" else "stocks"
        event[target_key] = {sym: safe_num(qty) for sym, qty in stocks.items()}
        events.append(event)

    events.sort(key=lambda e: e["date"])
    return events


transaction_events = build_transaction_events(holdings_hist, stock_codes, _dw)


# ── 7. TÍNH CÁC CHỈ SỐ HIỆU SUẤT ─────────────────────────────────────────────
df_pr["E0"] = df_pr["E1"].shift(-1)
df_pr.loc[df_pr.index[-1], "E0"] = 0

_dr_base           = df_pr["E0"] + df_pr["D"]
df_pr["DR"]        = (((df_pr["E1"] + df_pr["W"]) - _dr_base) / _dr_base.where(_dr_base != 0)).fillna(0.0)   # không có vốn -> DR = 0
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


# ── 7c. SHARPE & MAX DRAWDOWN — BẢN TỔNG LỊCH SỬ (mục 3) ─────────────────────
# Không group theo năm — chạy trên toàn bộ chuỗi DR/CR+1. Sharpe tổng dùng CAGR
# (không dùng CR thô) vì CR trải nhiều năm sẽ méo tỷ lệ so với rf (lãi suất/năm).
N_total   = len(df_pr)
cr_latest = df_pr.loc[df_pr.index[0], "CR"]  # index 0 = ngày mới nhất (df_pr sort giảm dần)

if N_total and (1 + cr_latest) > 0:
    cagr_total = (1 + cr_latest) ** (TRADING_DAYS / N_total) - 1
else:
    cagr_total = np.nan  # (1+CR) <= 0 sẽ ra số phức với số mũ lẻ — coi như không xác định

std_total = df_pr["DR"].std(ddof=0)
sharpe_total = (
    (cagr_total - RISK_FREE_RATE) / (std_total * np.sqrt(TRADING_DAYS))
    if std_total and pd.notna(cagr_total) else np.nan
)

cr1_all   = pd.to_numeric(df_pr["CR+1"], errors="coerce")[::-1]  # đảo về thời gian tăng dần
peaks_all = cr1_all.cummax()
max_dd_total = ((cr1_all - peaks_all) / peaks_all).min()

df_pr["Sharpe_Total"]      = sharpe_total
df_pr["MaxDrawdown_Total"] = max_dd_total


# ── 7d. RISK/MARKET METRICS: BETA, ALPHA, VOLATILITY, CORRELATION (mục 2) ───
rf_daily = RISK_FREE_RATE / TRADING_DAYS


def _risk_metrics(dr: pd.Series, dr_vni: pd.Series) -> dict:
    """Beta/Alpha/Volatility/Correlation trên 1 lát cắt DR/DR(VNI). Alpha và
    Volatility luôn annualize bằng TRADING_DAYS (252) — kể cả khi lát cắt là
    bản Tổng nhiều năm hay 1 năm có số ngày lẻ (chốt thiết kế: không đổi 252)."""
    dr     = pd.to_numeric(dr, errors="coerce")
    dr_vni = pd.to_numeric(dr_vni, errors="coerce")
    dr_mean, vni_mean = dr.mean(), dr_vni.mean()

    cov     = ((dr - dr_mean) * (dr_vni - vni_mean)).mean()  # Cov quần thể (ddof=0)
    var_vni = ((dr_vni - vni_mean) ** 2).mean()               # Var quần thể, đồng bộ STDEVP
    beta    = cov / var_vni if var_vni else np.nan

    alpha_daily = dr_mean - (rf_daily + beta * (vni_mean - rf_daily)) if pd.notna(beta) else np.nan
    alpha       = alpha_daily * TRADING_DAYS if pd.notna(alpha_daily) else np.nan

    volatility  = dr.std(ddof=0) * np.sqrt(TRADING_DAYS)
    correlation = dr.corr(dr_vni)

    return {"beta": beta, "alpha": alpha, "volatility": volatility, "correlation": correlation}


df_pr["Beta"]        = np.nan
df_pr["Alpha"]       = np.nan
df_pr["Volatility"]  = np.nan
df_pr["Correlation"] = np.nan

for _, idx in df_pr.groupby("Year_tmp").groups.items():
    idx = sorted(idx)
    m = _risk_metrics(df_pr.loc[idx, "DR"], df_pr.loc[idx, "DR(VNI)"])
    df_pr.loc[idx, "Beta"]        = m["beta"]
    df_pr.loc[idx, "Alpha"]       = m["alpha"]
    df_pr.loc[idx, "Volatility"]  = m["volatility"]
    df_pr.loc[idx, "Correlation"] = m["correlation"]

risk_total = _risk_metrics(df_pr["DR"], df_pr["DR(VNI)"])
df_pr["Beta_Total"]        = risk_total["beta"]
df_pr["Alpha_Total"]       = risk_total["alpha"]
df_pr["Volatility_Total"]  = risk_total["volatility"]
df_pr["Correlation_Total"] = risk_total["correlation"]


# ── 7e. XIRR (mục 5) — lãi suất nội bộ trên chuỗi cashflow thật (D/W) ────────
# Quy ước dấu: D = âm (tiền ra mua CP), W = dương (tiền về từ bán CP) — khớp mục 1,
# vì danh mục không giữ tiền mặt nên D/W trong holdings_history.csv chính là dòng tiền thật.
def _parsed_cashflows(cashflows_df: pd.DataFrame) -> list:
    parsed = []
    for _, row in cashflows_df.iterrows():
        try:
            d = datetime.strptime(str(row["date"]).strip(), "%d/%m/%Y").date()
        except ValueError:
            continue
        try:
            amt = float(row["value"])
        except (TypeError, ValueError):
            continue
        amt = -amt if str(row["type"]).strip().upper() == "D" else amt
        parsed.append((d, amt))
    return parsed


_cf_parsed = _parsed_cashflows(cashflows)

# XIRR tổng: toàn bộ D/W (holdings_history.csv) + 1 dòng cuối = E1 hiện tại (dương), ngày mới nhất
_latest_idx  = df_pr.index[0]
_latest_date = df_pr.loc[_latest_idx, "date"].date()
_latest_e1   = float(df_pr.loc[_latest_idx, "E1"])
_cfs_total   = _cf_parsed + [(_latest_date, _latest_e1)]

try:
    xirr_total = xirr(_cfs_total)
except Exception:
    xirr_total = None

# Lợi nhuận tổng (giá trị tuyệt đối, VNĐ) = E1 hiện tại - vốn ròng đã bỏ vào.
# Vì amt trong _cf_parsed đã mang dấu chuẩn (D = âm, W = dương, giống góc nhìn
# NĐT), tổng toàn bộ dòng tiền trong _cfs_total (kể cả E1 cuối) chính là lợi
# nhuận: E1 + sum(D âm) + sum(W dương) = E1 - (tổng D) + (tổng W).
profit_total = sum(amt for _, amt in _cfs_total)

# XIRR theo năm: E0 đầu năm (âm, giá trị E1 tại ngày giao dịch đầu tiên của năm) +
# D/W phát sinh trong năm (nguyên dấu, không double-count với E0/E1 vì đó chỉ là
# mốc định giá) + E1 cuối năm (dương, giá trị E1 tại ngày giao dịch cuối năm).
df_pr["XIRR_Year"]   = np.nan
df_pr["ProfitYear"]  = np.nan
for year, idx in df_pr.groupby("Year_tmp").groups.items():
    idx = sorted(idx)
    start_idx, end_idx = idx[-1], idx[0]  # df_pr giảm dần theo ngày: idx[-1]=sớm nhất, idx[0]=muộn nhất trong năm
    e0_date  = df_pr.loc[start_idx, "date"].date()
    e0_value = float(df_pr.loc[start_idx, "E1"])
    e1_date  = df_pr.loc[end_idx, "date"].date()
    e1_value = float(df_pr.loc[end_idx, "E1"])

    year_cfs = [(d, amt) for d, amt in _cf_parsed if d.year == year]
    # Năm đầu tiên: vốn ban đầu đã nằm trong D của ngày đầu, không cộng thêm E0 (tránh tính 2 lần)
    start_cf = [] if year == df_pr["Year_tmp"].min() else [(e0_date, -e0_value)]
    cfs = start_cf + year_cfs + [(e1_date, e1_value)]

    try:
        xirr_year = xirr(cfs)
    except Exception:
        xirr_year = None
    df_pr.loc[idx, "XIRR_Year"] = xirr_year

    # Lợi nhuận trong năm (giá trị tuyệt đối, VNĐ) = E1 cuối năm - E0 đầu năm -
    # cộng dồn cashflow phát sinh trong năm. Dùng đúng danh sách cfs ở trên
    # (chỉ khác dấu E0 đã đảo âm để tính XIRR) nên tổng cfs = lợi nhuận năm.
    df_pr.loc[idx, "ProfitYear"] = sum(amt for _, amt in cfs)


# ── 7f. ALLOCATION + POSITION CONTRIBUTION (mục 4) — chỉ tính ngày mới nhất ──
def compute_positions(df_pr: pd.DataFrame, holdings: pd.DataFrame, codes: list) -> dict:
    if df_pr.empty:
        return {"as_of_date": None, "items": []}

    latest_idx  = df_pr.index[0]  # df_pr sắp xếp giảm dần theo ngày → dòng đầu = mới nhất
    latest_date = df_pr.loc[latest_idx, "date"]
    year        = latest_date.year

    year_rows = df_pr[df_pr["Year_tmp"] == year]
    year_start_idx = year_rows["date"].idxmin() if not year_rows.empty else None

    def price_on(symbol, dt):
        """Giá đóng cửa của `symbol` đúng ngày `dt` (khớp theo price_history)."""
        if dt is None or pd.isna(dt):
            return None
        rows = df_pr.loc[df_pr["date"] == dt, symbol]
        if rows.empty:
            return None
        v = rows.iloc[0]
        return float(v) if pd.notna(v) else None

    def first_held_date(sym, only_year=None):
        """Ngày đầu tiên (trong lịch sử, hoặc trong năm only_year) mà mã được nắm giữ (qty > 0)."""
        mask = holdings[sym] > 0
        if only_year is not None:
            mask &= df_pr["Year_tmp"] == only_year
        dts = df_pr.loc[mask, "date"]
        return dts.min() if not dts.empty else None

    # Bước 1: qty/giá/giá trị từng mã đang có qty > 0 tại ngày mới nhất
    raw_positions = []
    total_value = 0.0
    for sym in codes:
        qty = float(holdings.loc[latest_idx, sym])
        if qty <= 0:
            continue
        price = float(df_pr.loc[latest_idx, sym])
        value = qty * price
        total_value += value
        raw_positions.append({"symbol": sym, "qty": qty, "price": price, "value": value})

    # Bước 2: weight + return_year/return_total + contribution
    items = []
    for p in raw_positions:
        sym, qty, price, value = p["symbol"], p["qty"], p["price"], p["value"]
        weight = value / total_value if total_value else None

        # return_year: giá mới nhất / giá ngày giao dịch đầu tiên của năm hiện tại - 1.
        # Nếu mã mới nắm giữ trong năm (chưa có tại ngày đầu năm) thì dùng ngày đầu
        # tiên mã đó xuất hiện trong năm thay cho đầu năm.
        held_at_year_start = year_start_idx is not None and holdings.loc[year_start_idx, sym] > 0
        if held_at_year_start:
            start_price_year = price_on(sym, df_pr.loc[year_start_idx, "date"])
        else:
            first_this_year = first_held_date(sym, only_year=year)
            if first_this_year is not None:
                start_price_year = price_on(sym, first_this_year)
            elif year_start_idx is not None:
                start_price_year = price_on(sym, df_pr.loc[year_start_idx, "date"])
            else:
                start_price_year = None
        return_year = (price / start_price_year - 1) if start_price_year else None

        # return_total: giá mới nhất / giá ngày đầu tiên nắm giữ (toàn bộ lịch sử của
        # holdings_history.csv) - 1.
        first_total = first_held_date(sym)
        start_price_total = price_on(sym, first_total) if first_total is not None else None
        return_total = (price / start_price_total - 1) if start_price_total else None

        contribution_year  = weight * return_year  if (weight is not None and return_year  is not None) else None
        contribution_total = weight * return_total if (weight is not None and return_total is not None) else None

        items.append({
            "symbol": sym,
            "sector": SECTOR_MAP.get(sym, SECTOR_DEFAULT),
            "qty":    safe_num(qty),
            "price":  safe_num(price),
            "value":  safe_num(value),
            "weight": safe_num(weight),
            "return_year":        safe_num(return_year),
            "return_total":       safe_num(return_total),
            "contribution_year":  safe_num(contribution_year),
            "contribution_total": safe_num(contribution_total),
        })

    return {"as_of_date": latest_date.strftime("%Y-%m-%d"), "items": items}


positions_data = compute_positions(df_pr, holdings, stock_codes)


# ── 8. XUẤT portfolio_history.csv (toàn bộ dữ liệu đã tính) ─────────────────
# File này là OUTPUT (derived), ghi đè hoàn toàn mỗi lần chạy — không phải nơi
# lưu trữ dữ liệu gốc nữa. Nguồn dữ liệu gốc thật sự là data/dnse/holdings_history.csv
# (từ DNSE) và price_history.csv (giá thô) — 2 file đó không được xoá.
df_pr["date"] = pd.to_datetime(df_pr["date"]).dt.strftime("%d/%m/%Y")
df_pr = df_pr.drop(columns=["date_str"], errors="ignore")

cols_to_drop = ["DR+1", "DR+1(VNI)", "CR+1", "CR+1(VNI)", "CR(VNI)",
                "MonthYear", "Cumulative_DR_M", "Year", "Year_tmp",
                "Cumulative_DR_Y", "Cumulative_DR_Y(VNI)"]
df_save = df_pr.drop(columns=[c for c in cols_to_drop if c in df_pr.columns])
PORTFOLIO_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
df_save.to_csv(PORTFOLIO_HISTORY_FILE, index=False)
print(f"✅ Đã lưu: {PORTFOLIO_HISTORY_FILE}")


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
    "date", *stock_codes, "VNINDEX",
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

apply_style_to_cols(ws, [*stock_codes, "VNINDEX", "E1", "W", "E0", "D"], int_acc, hm)
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
    "date": "C0FFC0",
    **{c: ("C0FFC0" if i % 2 == 0 else "CCFFFF") for i, c in enumerate(stock_codes)},
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
    "date": 10.6,
    **{c: 8.6 for c in stock_codes},
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
def export_json(df: pd.DataFrame, path: Path, events: list = None, positions: dict = None) -> None:
    records = df.copy()
    records["date"] = pd.to_datetime(records["date"], format="%d/%m/%Y")
    records = records.sort_values("date")
    records["date"] = records["date"].dt.strftime("%Y-%m-%d")

    keep_cols = [c for c in [
        "date", *stock_codes, "VNINDEX",
        "E1", "W", "E0", "D", "DR", "DR(VNI)", "CR", "CR(VNI)",
        "MR", "YR", "YR(VNI)", "Sharpe", "MaxDrawdown",
        "Sharpe_Total", "MaxDrawdown_Total",
        "Beta", "Alpha", "Volatility", "Correlation", "XIRR_Year", "ProfitYear",
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
        "sharpe_total": safe_num(latest.get("Sharpe_Total")),
        "max_drawdown_total": safe_num(latest.get("MaxDrawdown_Total")),
        "beta_total": safe_num(risk_total.get("beta")),
        "alpha_total": safe_num(risk_total.get("alpha")),
        "volatility_total": safe_num(risk_total.get("volatility")),
        "correlation_total": safe_num(risk_total.get("correlation")),
        "xirr_total": safe_num(xirr_total),
        "profit_total": safe_num(profit_total),
    }

    payload = {"summary": summary, "stock_codes": stock_codes, "history": history}
    if events is not None:
        payload["events"] = events
    if positions is not None:
        payload["positions"] = positions

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, allow_nan=False)
    print(f"✅ Đã xuất: {path}")


export_json(df_pr, JSON_FILE, events=transaction_events, positions=positions_data)
