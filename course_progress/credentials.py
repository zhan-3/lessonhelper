"""Windows-bound encrypted storage for optional WebVPN login credentials."""

from __future__ import annotations

import ctypes
import json
import os
from collections.abc import Callable
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import Path

CREDENTIALS_FILE_NAME = "webvpn-login.dpapi"
AUTH_STATE_FILE_NAME = "webvpn-auth-state.json"


@dataclass(frozen=True)
class LoginCredentials:
    username: str
    password: str


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _blob(value: bytes) -> tuple[_DataBlob, ctypes.Array]:
    buffer = ctypes.create_string_buffer(value)
    return (
        _DataBlob(len(value), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))),
        buffer,
    )


def _crypt_data(value: bytes, *, protect: bool) -> bytes:
    if os.name != "nt":
        raise RuntimeError("凭据加密仅支持 Windows DPAPI")
    source, source_buffer = _blob(value)
    output = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    function = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    if protect:
        ok = function(
            ctypes.byref(source),
            "HITWH WebVPN login",
            None,
            None,
            None,
            0x1,
            ctypes.byref(output),
        )
    else:
        ok = function(
            ctypes.byref(source),
            None,
            None,
            None,
            None,
            0x1,
            ctypes.byref(output),
        )
    del source_buffer  # keep the input buffer alive until the DPAPI call completes
    if not ok:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(output.pbData, output.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(output.pbData)


def protect_data(value: bytes) -> bytes:
    return _crypt_data(value, protect=True)


def unprotect_data(value: bytes) -> bytes:
    return _crypt_data(value, protect=False)


class CredentialStore:
    """Persist one login encrypted for the current Windows user account."""

    def __init__(
        self,
        path: Path,
        *,
        protect: Callable[[bytes], bytes] = protect_data,
        unprotect: Callable[[bytes], bytes] = unprotect_data,
    ):
        self.path = path
        self._protect = protect
        self._unprotect = unprotect

    def save(self, credentials: LoginCredentials) -> None:
        if not credentials.username or not credentials.password:
            raise ValueError("学号和密码不能为空")
        payload = json.dumps(
            {"username": credentials.username, "password": credentials.password},
            ensure_ascii=False,
        ).encode("utf-8")
        encrypted = self._protect(payload)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        temporary.write_bytes(encrypted)
        temporary.replace(self.path)

    def load(self) -> LoginCredentials | None:
        if not self.path.is_file():
            return None
        try:
            payload = json.loads(self._unprotect(self.path.read_bytes()).decode("utf-8"))
            username = str(payload["username"])
            password = str(payload["password"])
        except (OSError, UnicodeDecodeError, json.JSONDecodeError, KeyError, TypeError):
            raise RuntimeError("无法解密 WebVPN 登录凭据；请重新运行 configure-login")
        if not username or not password:
            raise RuntimeError("WebVPN 登录凭据不完整；请重新运行 configure-login")
        return LoginCredentials(username=username, password=password)


def credential_store(private_root: Path) -> CredentialStore:
    return CredentialStore(private_root.resolve() / CREDENTIALS_FILE_NAME)
