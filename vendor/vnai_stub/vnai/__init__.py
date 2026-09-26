"""
Stub thay thế cho package `vnai` (bị PyPI quarantine từ ~09/2026).

`vnai` là module analytics/license nội bộ của hệ sinh thái vnstock,
KHÔNG ảnh hưởng tới việc lấy dữ liệu chứng khoán. vnstock chỉ dùng nó để:
  - `optimize_execution(...)`: decorator (rate-limit/telemetry) bọc quanh
    các hàm gọi API — ở đây cho chạy thẳng hàm gốc, không làm gì thêm.
  - `accept_license_terms`, `setup_api_key`, `check_api_key_status`: chỉ
    được gọi trong các nhánh try/except khi đăng ký API key trả phí,
    không nằm trên đường chạy chính (Quote.history(), v.v.).

Khi PyPI gỡ quarantine cho `vnai`, có thể bỏ stub này và cài `vnai` thật
từ PyPI như bình thường.
"""

__version__ = "0.0.0"


def optimize_execution(*_args, **_kwargs):
    """No-op thay cho decorator gốc: trả nguyên hàm, không rate-limit/telemetry."""

    def _decorator(func):
        return func

    return _decorator


def accept_license_terms(*_args, **_kwargs):
    return True


def setup_api_key(*_args, **_kwargs):
    return True


def check_api_key_status(*_args, **_kwargs):
    return {"valid": True, "tier": "community"}


__all__ = [
    "optimize_execution",
    "accept_license_terms",
    "setup_api_key",
    "check_api_key_status",
]
