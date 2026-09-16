"""统一 HTTP 业务错误（§45 错误格式）。"""
from __future__ import annotations


class BackendHTTPError(Exception):
    """携带错误码的 HTTP 业务错误，由 main.py 统一转换为 {"error": {code, message}}。"""

    def __init__(self, status_code: int, code: str, message: str) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.code = code
        self.message = message
