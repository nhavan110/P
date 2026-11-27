# Portfolio Tracker

Tự động cập nhật giá trị danh mục đầu tư mỗi ngày qua GitHub Actions, xuất ra
Excel (`MyPortfolio.xlsx`), JSON cho dashboard web (`data.json`), và backup
lên Google Drive.

## Kiến trúc dữ liệu

```
[Máy local, IP VN]                                 [GitHub: repo P]
event_manual.py ─► event_manual.csv ─┐
                ─► positions.csv ────┼─ publish_to_github.py ─► data/dnse/*  ─► update_portfolio.py
                ─► holdings_history.csv ┘   (manifest.json đẩy cuối, kích hoạt workflow)   (giá từ API)

data/dnse/holdings_history.csv ─┐
data/price_history.csv ─────────┴─► update_portfolio.py ─┬─► data/portfolio_history.csv
                                                           ├─► data/data.json
                                                           └─► data/MyPortfolio.xlsx ──► Google Drive
```

P **không còn nhập tay**: mọi thứ về danh mục (mã, số lượng, nạp/rút) lấy từ DNSE; P chỉ lấy giá từ API.

| File | Vai trò | Được sửa tay? |
|---|---|---|
| `data/dnse/holdings_history.csv` | **Nguồn tính toán chính**: số lượng từng mã + D/W theo ngày (DNSE đẩy lên) | ❌ Không (ghi đè mỗi lần DNSE chạy) |
| `data/dnse/positions.csv` | Vị thế hiện tại của DNSE, chỉ để đối chiếu dòng mới nhất (✅ khớp / ⚠️ lệch) | ❌ Không |
| `data/dnse/event_manual.csv` | Sổ sự kiện DNSE, chỉ để tra cứu, P không dùng để tính | ❌ Không |
| `data/dnse/manifest.json` | Đẩy cuối cùng, đổi là kích hoạt workflow | ❌ Không |
| `data/price_history.csv` | Giá đóng cửa thô (các mã + VNINDEX). Mã mới sẽ được bù giá lịch sử từ API | ✅ Có (để sửa giá sai) |
| `data/portfolio_history.csv`, `data/data.json`, `data/MyPortfolio.xlsx` | **Output**, ghi đè mỗi lần chạy | ❌ Không |
| `data/archive/` | Dữ liệu giai đoạn cũ (2022 → trước khi dùng DNSE), P không đọc | — |

### `holdings_history.csv`

```
date,HPG,TCB,FPT,D,W
26/11/2026,500,400,600,0,0
23/11/2026,100,200,200,7015000,0
```

- Ngày `dd/mm/yyyy`. Chỉ có các ngày số lượng đổi hoặc có D/W, cộng dòng hôm nay; P tự điền các ngày giữa
  (lấy dòng gần nhất không muộn hơn ngày đó).
- Cột mã động: mọi cột trừ `date`, `D`, `W` là mã CK, không giới hạn số mã. Mã chưa có ngành trong `SECTOR_MAP` hiện "Khác".
- `D`/`W` = dòng tiền ròng của ngày (tiền mua − tiền bán đã gồm phí/thuế + phí lưu ký − cổ tức nhận).
  D/W rơi vào ngày nghỉ được dồn sang phiên giao dịch kế tiếp.
- Lịch sử và so sánh VNINDEX bắt đầu từ ngày đầu tiên của file này; lợi nhuận, Sharpe, XIRR tính lại từ đó.
- Nếu ngày đầu đã có sẵn cổ phiếu mà không có lệnh mua (mua trước `DNSE_FROM_DATE`), giá trị thị trường của số cổ phiếu đó
  được coi là vốn nạp ban đầu (D) của ngày đầu.
- Đổi số lượng mà không có tiền (D = W = 0) được coi là cổ phiếu thưởng/chia tách trên biểu đồ.
- E1 = Σ giá × số lượng (không còn margin).

## Các flow thường dùng

### 1. Giao dịch mới
Không làm gì ở P. Chạy DNSE ở máy local (`run.bat`); nó tự đẩy `data/dnse/*` rồi `manifest.json`,
workflow của P tự chạy và cập nhật dashboard.

### 2. Sửa 1 giá đóng cửa bị sai
Sửa trực tiếp `data/price_history.csv`, tìm đúng dòng theo ngày, sửa số trong đúng cột mã CK (hoặc VNINDEX).
Script chỉ fetch giá cho những ngày **sau** ngày mới nhất đang có trong file, nên giá bạn sửa không bị ghi đè lại.

## Vận hành GitHub Actions

- Workflow `.github/workflows/update.yml` chạy khi `data/dnse/manifest.json` đổi (DNSE vừa đẩy dữ liệu),
  và **1 lần/ngày lúc 18:00 giờ VN** (`0 11 * * *` UTC) để cập nhật giá; hoặc bấm **Run workflow** trên tab
  Actions để chạy tay bất cứ lúc nào.
- Thứ tự các bước: `update_portfolio.py` (tính toán + ghi các file trong
  `data/`) → `upload_to_drive.py` (đẩy `MyPortfolio.xlsx` lên Google Drive)
  → commit + push toàn bộ thay đổi trong `data/` về repo.
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

- **Không** sửa tay `holdings_history.csv`, `portfolio_history.csv`, `data.json`, `MyPortfolio.xlsx` — mất ngay lần chạy sau.
- Nếu `holdings_history.csv` chưa có mã nào (DNSE chưa có vị thế/giao dịch), script thoát sớm, không ghi output.
- Dòng cuối của `holdings_history.csv` lệch `positions.csv` sẽ in cảnh báo ⚠️ trong log Actions.
- Lỗi Google Drive 403 thường do: (1) chưa share quyền Editor folder cho
  service account, hoặc (2) chưa enable Google Drive API cho project đó
  trong Google Cloud Console.
