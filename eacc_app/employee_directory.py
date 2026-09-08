from __future__ import annotations

import ctypes
import json
from collections.abc import Mapping
from ctypes import wintypes
from pathlib import Path
from typing import Any


class EmployeeDirectoryError(RuntimeError):
    """Raised when the encrypted employee directory cannot be used safely."""


_HEADER = b"EACCMAIL1\x00"
_NAME_FIELDS = {"성명", "이름", "name", "employee_name", "employeename"}
_EMAIL_FIELDS = {"이메일", "이메일주소", "email", "email_address", "mail"}
_DEPARTMENT_FIELDS = {"부서", "부서명", "department", "team", "조직"}
_DPAPI_ENTROPY = b"E-AccAutoProcess|MailInfo|v1"


def default_employee_directory_path() -> Path:
    return Path(__file__).resolve().parent.parent / "config" / "mail_recipients.eacc"


def load_employee_names(path: Path | None = None) -> tuple[str, ...]:
    """Read only employee names from the encrypted mail-recipient file.

    Email addresses are intentionally not returned or logged.  The existing
    ``EACCMAIL1`` format stores a Windows DPAPI-protected UTF-8 JSON payload.
    """
    decoded = _load_directory_payload(path)
    names = _extract_names(decoded)
    if not names:
        raise EmployeeDirectoryError("직원 명단에서 성명 정보를 찾지 못했습니다.")
    return tuple(sorted(names))


def load_mail_recipients(path: Path | None = None) -> tuple[dict[str, str], ...]:
    """Return recipient records only inside the running process.

    Callers must never persist plaintext addresses to the ordinary activity log.
    """
    decoded = _load_directory_payload(path)
    records: list[dict[str, str]] = []
    for item in _walk_mappings(decoded):
        normalized = {str(key).replace(" ", "").casefold(): str(value).strip() for key, value in item.items()}
        name = next((normalized[key] for key in _NAME_FIELDS if normalized.get(key)), "")
        email = next((normalized[key] for key in _EMAIL_FIELDS if normalized.get(key)), "")
        department = next((normalized[key] for key in _DEPARTMENT_FIELDS if normalized.get(key)), "")
        if name and email:
            records.append({"name": name, "email": email, "department": department})
    if not records:
        raise EmployeeDirectoryError("직원 명단에서 이메일 정보를 찾지 못했습니다.")
    return tuple(records)


def _load_directory_payload(path: Path | None) -> Any:
    source = path or default_employee_directory_path()
    try:
        encrypted = source.read_bytes()
    except OSError as exc:
        raise EmployeeDirectoryError("직원 명단 암호화 파일을 읽을 수 없습니다.") from exc
    prefix_size = len(_HEADER)
    if not encrypted.startswith(_HEADER) or len(encrypted) <= prefix_size:
        raise EmployeeDirectoryError("직원 명단 암호화 파일 형식이 올바르지 않습니다.")
    try:
        payload = _unprotect_dpapi(encrypted[prefix_size:])
        decoded: Any = json.loads(payload.decode("utf-8"))
    except EmployeeDirectoryError:
        raise
    except Exception as exc:
        raise EmployeeDirectoryError("직원 명단 암호화 파일을 해독할 수 없습니다.") from exc

    return decoded


def _extract_names(value: Any) -> set[str]:
    names: set[str] = set()

    def visit(item: Any) -> None:
        if isinstance(item, Mapping):
            for key, child in item.items():
                normalized_key = str(key).replace(" ", "").casefold()
                if normalized_key in _NAME_FIELDS and isinstance(child, str):
                    cleaned = child.strip()
                    if cleaned:
                        names.add(cleaned)
                visit(child)
        elif isinstance(item, list):
            for child in item:
                visit(child)

    visit(value)
    return names


def _walk_mappings(value: Any):
    if isinstance(value, Mapping):
        yield value
        for child in value.values():
            yield from _walk_mappings(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_mappings(child)


def _unprotect_dpapi(ciphertext: bytes) -> bytes:
    if not hasattr(ctypes, "WinDLL"):
        raise EmployeeDirectoryError("Windows에서만 직원 명단을 해독할 수 있습니다.")

    class DataBlob(ctypes.Structure):
        _fields_ = [
            ("cbData", wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_char)),
        ]

    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    crypt32.CryptUnprotectData.argtypes = (
        ctypes.POINTER(DataBlob),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(DataBlob),
        wintypes.LPVOID,
        wintypes.LPVOID,
        wintypes.DWORD,
        ctypes.POINTER(DataBlob),
    )
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = (wintypes.HLOCAL,)
    kernel32.LocalFree.restype = wintypes.HLOCAL

    input_buffer = ctypes.create_string_buffer(ciphertext)
    entropy_buffer = ctypes.create_string_buffer(_DPAPI_ENTROPY)
    input_blob = DataBlob(len(ciphertext), ctypes.cast(input_buffer, ctypes.POINTER(ctypes.c_char)))
    entropy_blob = DataBlob(
        len(_DPAPI_ENTROPY), ctypes.cast(entropy_buffer, ctypes.POINTER(ctypes.c_char))
    )
    output_blob = DataBlob()
    description = wintypes.LPWSTR()
    if not crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        ctypes.byref(description),
        ctypes.byref(entropy_blob),
        None,
        None,
        0,
        ctypes.byref(output_blob),
    ):
        raise EmployeeDirectoryError("현재 Windows 계정에서 직원 명단을 해독할 수 없습니다.")
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(output_blob.pbData, wintypes.HLOCAL))
        if description:
            kernel32.LocalFree(description)
