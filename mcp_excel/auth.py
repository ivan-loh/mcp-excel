import os
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import Response
from starlette.requests import Request

from . import logging as log


class APIKeyMiddleware(BaseHTTPMiddleware):
    def __init__(self, app, api_key: str):
        super().__init__(app)
        self.api_key = api_key

    async def dispatch(self, request: Request, call_next):
        if not self.api_key:
            return await call_next(request)

        auth_header = request.headers.get("Authorization")

        if not auth_header:
            log.warn("auth_failed", reason="missing_header", path=request.url.path)
            return Response("Unauthorized: Missing Authorization header", status_code=401)

        if not auth_header.startswith("Bearer "):
            log.warn("auth_failed", reason="invalid_format", path=request.url.path)
            return Response("Unauthorized: Invalid Authorization format", status_code=401)

        token = auth_header[7:]

        if token != self.api_key:
            log.warn("auth_failed", reason="invalid_token", path=request.url.path)
            return Response("Unauthorized: Invalid API key", status_code=401)

        log.info("auth_success", path=request.url.path)
        return await call_next(request)


def get_api_key_from_env() -> str:
    return os.getenv("MCP_EXCEL_API_KEY", "")
