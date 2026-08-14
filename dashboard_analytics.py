# ── dashboard_analytics.py ───────────────────────────────────────────────────
# Tính toán các số liệu mở rộng cho dashboard (01, 03–07): allocation, P&L,
# position contribution, risk/market metrics, drawdown series, XIRR, và các
# điểm BUY/SELL để hiện trên chart. Tách riêng khỏi update_portfolio.py để cả
# pipeline chính (chạy trên GitHub Actions, có mạng) và script tái tạo JSON
# ngoại tuyến (offline, dùng dữ liệu cache sẵn) đều dùng chung một logic,
# tránh lệch kết quả giữa hai nơi.
from collections import deque

import numpy as np
import pandas as pd

TRADING_DAYS = 252
RISK_FREE_RATE = 0.04

SYMBOL_TO_COL = {
    "HPG": "b", "PNJ": "c", "TCB": "d", "MWG": "e",
    "MBB": "f", "FRT": "g", "FPT": "h",
}
COL_TO_SYMBOL = {v: k for k, v in SYMBOL_TO_COL.items()}
HOLDING_COLS = ["b", "c", "d", "e", "f", "g", "h", "margin"]

SECTOR_MAP = {
    "HPG": "Thép",
    "TCB": "Ngân hàng",
    "MBB": "Ngân hàng",
    "FPT": "Công nghệ",
    "PNJ": "Bán lẻ",
    "FRT": "Bán lẻ",
    "MWG": "Bán lẻ",
}

STOCK_CODES = ["HPG", "TCB", "FPT", "PNJ", "FRT", "MWG", "MBB"]


def safe_num(x):
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    if not np.isfinite(v):
        return None
    return v


# ── HOLDINGS (số lượng CP nắm giữ theo từng ngày) ───────────────────────────
def get_holdings(dates: pd.Series, transactions: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame(0.0, index=dates.index, columns=HOLDING_COLS)
    if transactions.empty:
        return result
    running = {c: 0.0 for c in HOLDING_COLS}
    for _, row in transactions.iterrows():
        if row["symbol"] == "MARGIN":
            col = "margin"
            running[col] = float(row["quantity"])
        else:
            col = SYMBOL_TO_COL[row["symbol"]]
            delta = row["quantity"] if row["action"] == "BUY" else -row["quantity"]
            running[col] += float(delta)
        result.loc[dates >= row["date"], col] = running[col]
    return result


# ── FIFO: khớp lệnh SELL với các lô BUY trước đó để tính P&L thực hiện ──────
def run_fifo(transactions: pd.DataFrame, price_lookup: dict):
    """transactions: date (Timestamp), symbol, action, quantity — đã sort tăng dần.
    price_lookup: dict[(date_str_ddmmyyyy, symbol)] -> giá đóng cửa ngày đó.
    Trả về: (sell_records, remaining_lots)
      sell_records: list các dict cho mỗi lệnh SELL {date, symbol, quantity,
                    price, value, cost_basis, realized_pnl}
      remaining_lots: dict[symbol] -> list các lô còn lại [qty, cost_price]
    """
    lots = {sym: deque() for sym in STOCK_CODES}
    sell_records = []
    buy_records = []

    for _, row in transactions.iterrows():
        sym = row["symbol"]
        if sym == "MARGIN":
            continue
        d_str = row["date"].strftime("%d/%m/%Y")
        price = price_lookup.get((d_str, sym))
        if price is None or price == 0:
            continue
        qty = float(row["quantity"])
        value = price * qty

        if row["action"] == "BUY":
            lots[sym].append([qty, price])
            buy_records.append({
                "date": row["date"].strftime("%Y-%m-%d"), "symbol": sym,
                "action": "BUY", "quantity": qty, "price": price, "value": value,
            })
        else:  # SELL
            remaining = qty
            cost_basis = 0.0
            while remaining > 1e-9 and lots[sym]:
                lot_qty, lot_price = lots[sym][0]
                matched = min(lot_qty, remaining)
                cost_basis += matched * lot_price
                lot_qty -= matched
                remaining -= matched
                if lot_qty <= 1e-9:
                    lots[sym].popleft()
                else:
                    lots[sym][0][0] = lot_qty
            realized_pnl = value - cost_basis
            sell_records.append({
                "date": row["date"].strftime("%Y-%m-%d"), "symbol": sym,
                "action": "SELL", "quantity": qty, "price": price, "value": value,
                "cost_basis": cost_basis, "realized_pnl": realized_pnl,
            })

    return buy_records, sell_records, lots


# ── XIRR (Newton's method, có fallback bisection) ───────────────────────────
def xirr(cashflows):
    """cashflows: list of (date: pd.Timestamp, amount: float). Trả về tỷ suất
    sinh lời năm hoá (XIRR) hoặc None nếu không hội tụ."""
    if len(cashflows) < 2:
        return None
    t0 = cashflows[0][0]
    days = np.array([(d - t0).days / 365.0 for d, _ in cashflows])
    amounts = np.array([a for _, a in cashflows])

    def npv(rate):
        return np.sum(amounts / (1.0 + rate) ** days)

    def dnpv(rate):
        return np.sum(-days * amounts / (1.0 + rate) ** (days + 1))

    rate = 0.1
    for _ in range(100):
        f = npv(rate)
        fp = dnpv(rate)
        if abs(fp) < 1e-12:
            break
        new_rate = rate - f / fp
        if not np.isfinite(new_rate) or new_rate <= -0.999:
            break
        if abs(new_rate - rate) < 1e-8:
            return new_rate
        rate = new_rate

    # Fallback: bisection trong khoảng hợp lý
    lo, hi = -0.9999, 10.0
    f_lo, f_hi = npv(lo), npv(hi)
    if f_lo * f_hi > 0:
        return rate if np.isfinite(rate) else None
    for _ in range(200):
        mid = (lo + hi) / 2
        f_mid = npv(mid)
        if abs(f_mid) < 1e-6:
            return mid
        if f_lo * f_mid < 0:
            hi = mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2


# ── Chỉ số rủi ro/thị trường (Alpha, Beta, Volatility, Correlation, Sharpe) ──
def risk_metrics(dr: pd.Series, dr_vni: pd.Series, period_return: float, n_days: int):
    dr = dr.astype(float).values
    dr_vni = dr_vni.astype(float).values
    out = {"alpha": None, "beta": None, "volatility": None, "correlation": None, "sharpe": None}
    if len(dr) < 2:
        return out

    var_vni = np.var(dr_vni, ddof=0)
    beta = np.cov(dr, dr_vni, ddof=0)[0, 1] / var_vni if var_vni else None
    vol_annual = float(np.std(dr, ddof=0) * np.sqrt(TRADING_DAYS))
    if beta is not None:
        alpha_annual = float((np.mean(dr) - beta * np.mean(dr_vni)) * TRADING_DAYS)
    else:
        alpha_annual = None
    std_vni = np.std(dr_vni, ddof=0)
    std_dr = np.std(dr, ddof=0)
    corr = float(np.corrcoef(dr, dr_vni)[0, 1]) if std_vni and std_dr else None

    annual_return = (1.0 + period_return) ** (TRADING_DAYS / n_days) - 1.0 if n_days > 0 else None
    sharpe = ((annual_return - RISK_FREE_RATE) / vol_annual) if (annual_return is not None and vol_annual) else None

    out["alpha"] = safe_num(alpha_annual)
    out["beta"] = safe_num(beta)
    out["volatility"] = safe_num(vol_annual)
    out["correlation"] = safe_num(corr)
    out["sharpe"] = safe_num(sharpe)
    return out


# ── HÀM CHÍNH ─────────────────────────────────────────────────────────────
def compute_dashboard_extras(df_hist: pd.DataFrame, transactions: pd.DataFrame) -> dict:
    """df_hist: DataFrame TĂNG DẦN theo ngày, có cột date(Timestamp), HPG..MBB,
    VNINDEX, E1, W, D, DR, DR(VNI).
    transactions: DataFrame TĂNG DẦN theo ngày, cột date(Timestamp), symbol,
    action, quantity (đã load từ transactions.csv, symbol có thể là MARGIN).
    """
    df = df_hist.sort_values("date").reset_index(drop=True)
    dates = df["date"]

    holdings = get_holdings(dates, transactions)
    latest = df.iloc[-1]
    latest_holdings = holdings.iloc[-1]

    # price_lookup cho FIFO
    price_lookup = {}
    for _, row in df.iterrows():
        d_str = row["date"].strftime("%d/%m/%Y")
        for sym in STOCK_CODES:
            price_lookup[(d_str, sym)] = row.get(sym, 0.0)

    tx_stock = transactions[transactions["symbol"] != "MARGIN"].copy()
    buy_records, sell_records, remaining_lots = run_fifo(tx_stock, price_lookup)

    # ── POSITIONS hiện tại (giữ lô còn lại, giá vốn bình quân) ──────────────
    positions = []
    total_mv = 0.0
    for col in ["b", "c", "d", "e", "f", "g", "h"]:
        sym = COL_TO_SYMBOL[col]
        qty = float(latest_holdings[col])
        if qty <= 1e-9:
            continue
        price = float(latest.get(sym, 0.0))
        mv = qty * price
        total_mv += mv
        lots = remaining_lots.get(sym, deque())
        lot_qty_sum = sum(l[0] for l in lots)
        cost_sum = sum(l[0] * l[1] for l in lots)
        avg_cost = (cost_sum / lot_qty_sum) if lot_qty_sum > 1e-9 else price
        positions.append({
            "symbol": sym, "quantity": qty, "avg_cost": safe_num(avg_cost),
            "current_price": safe_num(price), "market_value": safe_num(mv),
            "sector": SECTOR_MAP.get(sym, "Khác"),
        })

    for p in positions:
        p["weight"] = safe_num(p["market_value"] / total_mv) if total_mv else None
        ret = (p["current_price"] / p["avg_cost"] - 1.0) if p["avg_cost"] else None
        p["return"] = safe_num(ret)
        p["contribution"] = safe_num(p["weight"] * ret) if (p["weight"] is not None and ret is not None) else None

    positions.sort(key=lambda p: (p["weight"] or 0), reverse=True)

    sector_totals = {}
    for p in positions:
        sector_totals[p["sector"]] = sector_totals.get(p["sector"], 0.0) + (p["market_value"] or 0.0)
    sector_allocation = [
        {"sector": s, "weight": safe_num(v / total_mv) if total_mv else None, "market_value": safe_num(v)}
        for s, v in sorted(sector_totals.items(), key=lambda kv: kv[1], reverse=True)
    ]

    position_contribution = [
        {"symbol": p["symbol"], "weight": p["weight"], "return": p["return"], "contribution": p["contribution"]}
        for p in positions
    ]

    # ── Markers BUY/SELL cho chart (giá, khối lượng, giá trị, realized P&L) ─
    markers = []
    for r in buy_records:
        markers.append({"date": r["date"], "symbol": r["symbol"], "action": "BUY",
                         "price": safe_num(r["price"]), "quantity": safe_num(r["quantity"]),
                         "value": safe_num(r["value"])})
    for r in sell_records:
        markers.append({"date": r["date"], "symbol": r["symbol"], "action": "SELL",
                         "price": safe_num(r["price"]), "quantity": safe_num(r["quantity"]),
                         "value": safe_num(r["value"]), "realized_pnl": safe_num(r["realized_pnl"])})
    markers.sort(key=lambda m: m["date"])

    # ── Drawdown series (LUÔN trên toàn bộ lịch sử) ─────────────────────────
    equity = (df["DR"].astype(float) + 1.0).cumprod()
    peak = equity.cummax()
    dd = (equity - peak) / peak
    drawdown_series = [
        {"date": d.strftime("%Y-%m-%d"), "drawdown": safe_num(v)}
        for d, v in zip(df["date"], dd)
    ]
    max_drawdown_all = safe_num(dd.min())
    current_drawdown = safe_num(dd.iloc[-1])

    # ── XIRR (toàn bộ lịch sử) ───────────────────────────────────────────────
    cf = []
    for _, row in df.iterrows():
        if row.get("D", 0):
            cf.append((row["date"], -float(row["D"])))
        if row.get("W", 0):
            cf.append((row["date"], float(row["W"])))
    cf.append((df["date"].iloc[-1], float(latest["E1"])))
    xirr_all = safe_num(xirr(cf)) if len(cf) >= 2 else None

    # ── PERIODS: "ALL" + từng năm dương lịch ────────────────────────────────
    df["year"] = df["date"].dt.year
    net_invested_cum = (df["D"].astype(float) - df["W"].astype(float)).cumsum()

    # P&L theo CỔ PHIẾU (không tính margin) — snapshot "cộng dồn tới ngày X":
    # unrealized_asof(X) = giá trị thị trường các lô còn lại tính theo giá ngày
    # X trừ giá vốn các lô đó; realized_asof(X) = tổng lãi/lỗ đã chốt (FIFO)
    # của các lệnh SELL có ngày <= X. P&L một kỳ = (realized+unrealized cuối kỳ)
    # − (realized+unrealized đầu kỳ) → tách đúng phần lãi/lỗ phát sinh trong kỳ,
    # dù đến từ việc chốt lời hay từ biến động giá của các mã đang nắm giữ.
    price_by_date = {row["date"]: row for _, row in df.iterrows()}

    def snapshot_asof(boundary_date):
        if boundary_date is None:
            return 0.0, 0.0  # realized_cum, unrealized_cum
        tx_upto = tx_stock[tx_stock["date"] <= boundary_date]
        _, sells_upto, lots_upto = run_fifo(tx_upto, price_lookup)
        realized_cum = sum(r["realized_pnl"] for r in sells_upto)
        row = price_by_date.get(boundary_date)
        unrealized_cum = 0.0
        if row is not None:
            for sym, dq in lots_upto.items():
                qty = sum(l[0] for l in dq)
                cost = sum(l[0] * l[1] for l in dq)
                if qty > 1e-9:
                    unrealized_cum += qty * float(row.get(sym, 0.0)) - cost
        return realized_cum, unrealized_cum

    snap_cache = {}

    def get_snapshot(boundary_date):
        if boundary_date not in snap_cache:
            snap_cache[boundary_date] = snapshot_asof(boundary_date)
        return snap_cache[boundary_date]

    periods = {}
    years = sorted(df["year"].unique().tolist())

    def build_period(sub: pd.DataFrame, label: str, start_date, end_date,
                      start_value: float, net_invested_period: float):
        n_days = len(sub)
        cr_period = float((sub["DR"].astype(float) + 1.0).prod() - 1.0)
        cr_vni_period = float((sub["DR(VNI)"].astype(float) + 1.0).prod() - 1.0)
        end_value = float(sub["E1"].iloc[-1])
        rm = risk_metrics(sub["DR"], sub["DR(VNI)"], cr_period, n_days)

        realized_start, unrealized_start = get_snapshot(start_date)
        realized_end, unrealized_end = get_snapshot(end_date)
        pnl_realized = realized_end - realized_start
        pnl_unrealized = unrealized_end - unrealized_start
        pnl_total = pnl_realized + pnl_unrealized

        # XIRR riêng cho kỳ: coi giá trị đầu kỳ là một khoản "nạp" tại ngày đầu kỳ
        cf_p = []
        if start_value:
            cf_p.append((sub["date"].iloc[0], -start_value))
        for _, row in sub.iterrows():
            if row.get("D", 0):
                cf_p.append((row["date"], -float(row["D"])))
            if row.get("W", 0):
                cf_p.append((row["date"], float(row["W"])))
        cf_p.append((sub["date"].iloc[-1], end_value))
        xirr_p = safe_num(xirr(cf_p)) if len(cf_p) >= 2 else None

        return {
            "label": label,
            "portfolio_value": safe_num(end_value),
            "invested_capital": safe_num(net_invested_cum.loc[sub.index[-1]]),
            "pnl_realized": safe_num(pnl_realized),
            "pnl_unrealized": safe_num(pnl_unrealized),
            "pnl_total": safe_num(pnl_total),
            "return_pct": safe_num(cr_period),
            "vni_return_pct": safe_num(cr_vni_period),
            "xirr": xirr_p,
            "alpha": rm["alpha"], "beta": rm["beta"], "volatility": rm["volatility"],
            "correlation": rm["correlation"], "sharpe": rm["sharpe"],
            "max_drawdown": max_drawdown_all,
        }

    periods["ALL"] = build_period(df, "Tổng lịch sử", None, df["date"].iloc[-1], 0.0,
                                   float(net_invested_cum.iloc[-1]))

    for y in years:
        sub = df[df["year"] == y]
        prev = df[df["date"] < sub["date"].iloc[0]]
        start_value = float(prev["E1"].iloc[-1]) if len(prev) else 0.0
        start_date = prev["date"].iloc[-1] if len(prev) else None
        end_date = sub["date"].iloc[-1]
        periods[str(y)] = build_period(sub, str(y), start_date, end_date, start_value,
                                        float(net_invested_cum.loc[sub.index[-1]] - (net_invested_cum.loc[prev.index[-1]] if len(prev) else 0.0)))

    return {
        "positions": positions,
        "sector_allocation": sector_allocation,
        "position_contribution": position_contribution,
        "transactions_markers": markers,
        "drawdown_series": drawdown_series,
        "max_drawdown_all": max_drawdown_all,
        "current_drawdown": current_drawdown,
        "xirr_all": xirr_all,
        "periods": periods,
        "available_years": [str(y) for y in years],
    }
