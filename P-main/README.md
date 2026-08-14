# Portfolio Tracker

Tự động cập nhật giá trị danh mục đầu tư mỗi ngày qua GitHub Actions, xuất ra
Excel (`MyPortfolio.xlsx`), JSON cho dashboard web (`data.json`), và backup
lên Google Drive.

## Kiến trúc dữ liệu

```
data/transactions.csv  ──┐
data/price_history.csv ──┼──►  update_portfolio.py  ──┬──► data/portfolio_history.csv
data/cashflows.csv     ──┘         │                   ├──► data/data.json
                                    │                   └──► data/MyPortfolio.xlsx ──► Google Drive
                                    ▼
                          dashboard_analytics.py
                   (positions, P&L, allocation, risk metrics,
                    drawdown, XIRR, BUY/SELL markers cho index.html)
```

`dashboard_analytics.py` chứa toàn bộ logic tính cho 6 mục còn lại của
dashboard (Overview mở rộng, P&L, Allocation, Position Contribution,
Drawdown, Risk/Market Metrics) — được `update_portfolio.py` gọi ở bước xuất
`data.json`. Nếu cần tái tạo `data.json` mà không có mạng (không fetch giá
mới), dùng `python3 build_offline_preview.py` — script này đọc lại
`data/portfolio_history.csv` + `data/transactions.csv` đã cache sẵn, dùng
chung `dashboard_analytics.py`, chỉ khác nguồn giá đầu vào.

Nguyên tắc quan trọng nhất: **chỉ 3 file bên trái là dữ liệu gốc (source of
truth), 3 file bên phải chỉ là kết quả tính ra.**

| File | Vai trò | Được sửa tay? | Cách cập nhật |
|---|---|---|---|
| `data/transactions.csv` | Sổ mua/bán cổ phiếu + thay đổi margin | ✅ Có | Thêm 1 dòng mỗi khi giao dịch |
| `data/price_history.csv` | Giá đóng cửa thô (7 mã CK + VNINDEX) theo ngày | ✅ Có (để sửa giá sai) | Script tự fetch & nối thêm mỗi lần chạy |
| `data/cashflows.csv` | Sổ nạp/rút tiền (W/D), tích luỹ dần | ✅ Có (để sửa W/D sai) | Tự động gộp từ `manual_entries.csv` |
| `manual_entries.csv` | Inbox tạm để nhập W/D mới | ✅ Có | Thêm dòng, script tự xoá sau khi xử lý |
| `data/portfolio_history.csv` | **Output** — toàn bộ số liệu đã tính (E1, DR, CR, YR...) | ❌ Không | Ghi đè hoàn toàn mỗi lần chạy |
| `data/data.json` | **Output** — dữ liệu cho dashboard web (`index.html`) | ❌ Không | Ghi đè hoàn toàn mỗi lần chạy |
| `data/MyPortfolio.xlsx` | **Output** — bản Excel định dạng đẹp, upload lên Drive | ❌ Không | Ghi đè hoàn toàn mỗi lần chạy |

⚠️ Sửa tay vào `portfolio_history.csv` hoặc `data.json` sẽ **mất ngay ở lần
chạy kế tiếp** vì script không đọc lại 2 file này, chỉ ghi ra.

## Các flow thường dùng

### 1. Thêm 1 giao dịch mua/bán cổ phiếu

Sửa `data/transactions.csv` trên GitHub web (icon bút chì), thêm 1 dòng vào
cuối file theo đúng cột:

```
date,symbol,action,quantity,note
14/08/2026,PNJ,BUY,100,
```

- `date`: dd/mm/yyyy, ngày khớp lệnh
- `symbol`: HPG, PNJ, TCB, MWG, MBB, FRT, hoặc FPT
- `action`: `BUY` hoặc `SELL`
- `quantity`: số lượng cổ phiếu (số dương)
- `note`: ghi chú tự do, để trống vẫn phải có dấu phẩy phía trước

Holdings từ ngày đó trở đi tự tính lại đúng — không cần sửa gì khác.

### 2. Thay đổi số dư margin (vay ký quỹ)

Cũng trong `data/transactions.csv`, dùng symbol `MARGIN` với action `SET`
(vì margin là **số dư nợ hiện tại**, không phải số lượng cộng dồn như cổ
phiếu):

```
14/08/2026,MARGIN,SET,15000000,
```

Nghĩa là "từ ngày 14/08/2026, số dư margin = 15,000,000đ" cho tới khi có
dòng `MARGIN,SET` tiếp theo.

### 3. Nạp / rút tiền (W / D)

Sửa `manual_entries.csv`, thêm dòng:

```
date,type,value
14/08/2026,D,10000000
```

`type` là `D` (nạp/deposit) hoặc `W` (rút/withdraw). Lần chạy script kế tiếp
sẽ tự gộp dòng này vào `data/cashflows.csv` (sổ tích luỹ vĩnh viễn) rồi tự
xoá dòng đã xử lý khỏi `manual_entries.csv`. Nếu ngày đó chưa có trong
`price_history.csv` (chưa phải ngày giao dịch đã có giá), dòng sẽ bị bỏ qua
và in cảnh báo trong log Actions.

Nếu 1 ngày có nhiều lần nạp/rút, `cashflows.csv` cho phép nhiều dòng cùng
ngày cùng loại — chúng được **cộng dồn** khi tính toán, không ghi đè lẫn
nhau.

### 4. Sửa 1 giá đóng cửa bị sai (do API lỗi, hoặc số liệu cũ sai)

Sửa trực tiếp `data/price_history.csv`, tìm đúng dòng theo ngày, sửa số
trong đúng cột mã CK (hoặc VNINDEX). Vì script chỉ fetch giá cho những ngày
**sau** ngày mới nhất đang có trong file (`date > last_update`), giá trị bạn
sửa cho một ngày đã tồn tại sẽ **không bao giờ bị ghi đè lại** bởi lần chạy
sau, kể cả khi chạy lại nhiều lần trong cùng ngày.

### 5. Sửa 1 giao dịch hoặc 1 khoản W/D đã nhập sai trong quá khứ

Tìm đúng dòng trong `data/transactions.csv` hoặc `data/cashflows.csv`, sửa
trực tiếp số liệu (quantity / value) hoặc ngày, commit. Toàn bộ số liệu tính
toán liên quan (holdings, E1, DR, CR...) sẽ tự tính lại đúng ở lần chạy kế
tiếp.

## Vận hành GitHub Actions

- Workflow `.github/workflows/update.yml` chạy tự động **1 lần/ngày lúc
  18:00 giờ VN** (`0 11 * * *` UTC), hoặc bấm **Run workflow** trên tab
  Actions để chạy tay bất cứ lúc nào.
- Thứ tự các bước: `update_portfolio.py` (tính toán + ghi các file trong
  `data/`) → `upload_to_drive.py` (đẩy `MyPortfolio.xlsx` lên Google Drive)
  → commit + push toàn bộ thay đổi trong `data/` và `manual_entries.csv` về
  repo.
- Cần 2 secret trong repo Settings → Secrets → Actions:
  - `GOOGLE_SERVICE_ACCOUNT_JSON`: toàn bộ nội dung file JSON service
    account (copy nguyên khối, kể cả dấu `{ }`)
  - `GOOGLE_DRIVE_FOLDER_ID`: chỉ phần ID trong URL folder Drive, không phải
    cả link
  - Folder Drive đích phải được **share quyền Editor** cho email của service
    account (`xxx@xxx.iam.gserviceaccount.com`, xem trong file JSON), và
    **Google Drive API phải được enable** trong Google Cloud Console cho
    project chứa service account đó.

## Lưu ý / bẫy hay gặp

- **Không** sửa tay `portfolio_history.csv`, `data.json`, `MyPortfolio.xlsx`
  — mất ngay lần chạy sau.
- **Không** sửa `update_portfolio.py` để thêm giao dịch nữa — đó là lỗi
  thiết kế cũ đã bỏ. Giờ mọi giao dịch nằm trong `data/transactions.csv`.
- Định dạng ngày **luôn** là `dd/mm/yyyy` ở mọi file CSV nhập tay
  (`transactions.csv`, `cashflows.csv`, `manual_entries.csv`).
- Nếu gõ sai `symbol` hoặc `action` trong `transactions.csv`, script sẽ báo
  lỗi rõ ràng (`ValueError`, dừng chạy) thay vì âm thầm tính sai — Actions sẽ
  đỏ, xem log để biết chính xác dòng nào sai.
- Nếu `manual_entries.csv` có ngày không tồn tại trong `price_history.csv`
  (ví dụ ngày nghỉ giao dịch, hoặc gõ sai ngày), dòng đó bị bỏ qua và in
  cảnh báo — không raise lỗi, không dừng workflow.
- Lỗi Google Drive 403 thường do: (1) chưa share quyền Editor folder cho
  service account, hoặc (2) chưa enable Google Drive API cho project đó
  trong Google Cloud Console.
