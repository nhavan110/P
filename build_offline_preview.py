import json
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd

from dashboard_analytics import compute_dashboard_extras, safe_num, STOCK_CODES

VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")
BASE_DIR = Path(__file__).resolve().parent
PORTFOLIO_HISTORY_FILE = BASE_DIR / "data" / "portfolio_history.csv"
TRANSACTIONS_FILE = BASE_DIR / "data" / "transactions.csv"
JSON_FILE = BASE_DIR / "data" / "data.json"

# ── 1. Nạp portfolio_history.csv (đã có đủ E1/DR/DR(VNI)/CR... từ lần chạy trước) ──
df = pd.read_csv(PORTFOLIO_HISTORY_FILE)
df["date"] = pd.to_datetime(df["date"], format="%d/%m/%Y")
df = df.sort_values("date").reset_index(drop=True)

# CR(VNI) không có sẵn trong portfolio_history.csv (chỉ có DR(VNI)) -> tính lại
df["CR(VNI)"] = (df["DR(VNI)"].astype(float) + 1.0)[::-1].cumprod()[::-1] - 1.0

# ── 2. Nạp transactions.csv ──────────────────────────────────────────────────
tx = pd.read_csv(TRANSACTIONS_FILE, dtype={"note": str})
tx["date"] = pd.to_datetime(tx["date"], format="%d/%m/%Y")
tx = tx.sort_values("date").reset_index(drop=True)

# ── 3. Tính các số liệu mở rộng ──────────────────────────────────────────────
extras = compute_dashboard_extras(df, tx)

# ── 4. History cơ bản (giữ nguyên format cũ để không phá code hiện có) ───────
keep_cols = [c for c in [
    "date", "HPG", "TCB", "FPT", "PNJ", "FRT", "MWG", "MBB", "VNINDEX",
    "E1", "W", "E0", "D", "DR", "DR(VNI)", "CR", "CR(VNI)",
    "MR", "YR", "YR(VNI)", "Sharpe", "MaxDrawdown",
] if c in df.columns]

history = []
for _, row in df[keep_cols].iterrows():
    rec = {"date": row["date"].strftime("%Y-%m-%d")}
    for c in keep_cols:
        if c == "date":
            continue
        rec[c] = safe_num(row[c])
    history.append(rec)

latest = df.iloc[-1]
summary = {
    "last_updated": datetime.now(VN_TZ).strftime("%Y-%m-%d %H:%M:%S"),
    "portfolio_value": safe_num(latest.get("E1")),
    "cumulative_return": safe_num(df.iloc[-1]["CR"] if "CR" in df.columns else None),
    "cumulative_return_vni": safe_num(df.iloc[-1]["CR(VNI)"]),
    "sharpe_ratio": safe_num(latest.get("Sharpe")),
    "max_drawdown": extras["max_drawdown_all"],
}

out = {
    "summary": summary,
    "history": history,
    "positions": extras["positions"],
    "sector_allocation": extras["sector_allocation"],
    "position_contribution": extras["position_contribution"],
    "transactions_markers": extras["transactions_markers"],
    "drawdown_series": extras["drawdown_series"],
    "current_drawdown": extras["current_drawdown"],
    "xirr_all": extras["xirr_all"],
    "periods": extras["periods"],
    "available_years": extras["available_years"],
}

JSON_FILE.parent.mkdir(parents=True, exist_ok=True)
with open(JSON_FILE, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2, allow_nan=False)

print("OK, đã ghi", JSON_FILE)
print("positions:", len(extras["positions"]))
print("sector_allocation:", extras["sector_allocation"])
print("periods keys:", list(extras["periods"].keys()))
print("ALL period:", json.dumps(extras["periods"]["ALL"], ensure_ascii=False, indent=2))
