from __future__ import annotations

import re
from dataclasses import dataclass
from html import unescape
from http.cookiejar import CookieJar
from threading import Event, Lock, Thread
from urllib.parse import urlencode
from urllib.request import HTTPCookieProcessor, Request, build_opener


BIZNO_URL = "https://bizno.net/"


class MerchantLookupError(RuntimeError):
    """Raised when BizNo cannot return one reliable business name."""


@dataclass(frozen=True, slots=True)
class BusinessNumberExtraction:
    candidates: tuple[str, ...]
    reason: str = ""


def is_pg_business_type(business_type: str) -> bool:
    """Return true for every current and future label containing ``PG일반``."""
    return "PG일반" in (business_type or "")


def is_valid_business_number(value: str) -> bool:
    """Validate a Korean 10-digit business registration number checksum."""
    digits = re.sub(r"\D", "", value)
    if len(digits) != 10:
        return False
    weights = (1, 3, 7, 1, 3, 7, 1, 3, 5)
    total = sum(int(digit) * weight for digit, weight in zip(digits[:9], weights))
    total += (int(digits[8]) * 5) // 10
    return (10 - (total % 10)) % 10 == int(digits[9])


def extract_business_numbers(text: str) -> BusinessNumberExtraction:
    """Extract only checksum-valid business numbers from local OCR text.

    OCR often inserts spaces or hyphens into the conventional ``000-00-00000``
    representation.  This deliberately accepts those separators but does not
    guess missing digits: a candidate is usable only after checksum validation.
    """
    normalized = str(text or "").translate(str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "|": "1"}))
    patterns = (
        r"(?<!\d)(\d{3}\s*[-‐‑–—]?\s*\d{2}\s*[-‐‑–—]?\s*\d{5})(?!\d)",
        r"(?<!\d)(\d(?:[\d\s-]{8,16})\d)(?!\d)",
    )
    candidates: list[str] = []
    for pattern in patterns:
        for raw in re.findall(pattern, normalized):
            digits = re.sub(r"\D", "", raw)
            if is_valid_business_number(digits) and digits not in candidates:
                candidates.append(digits)
    if not candidates:
        return BusinessNumberExtraction((), "영수증 OCR에서 체크섬이 맞는 사업자번호를 찾지 못했습니다.")
    if len(candidates) > 1:
        return BusinessNumberExtraction(tuple(candidates), "체크섬이 맞는 사업자번호가 여러 개입니다.")
    return BusinessNumberExtraction(tuple(candidates))


def extract_business_name_from_text(text: str) -> str:
    """Conservatively read a business name from a labelled result text."""
    for line in str(text or "").splitlines():
        compact = " ".join(line.split())
        match = re.search(
            r"(?:상호명|사업자명|사업자\s*상호|회사명)\s*[:：]?\s*(.+)$",
            compact,
        )
        if not match:
            continue
        name = match.group(1).strip(" -:：")
        # A result label must be followed by a plausible name, not another label.
        if 2 <= len(name) <= 80 and not re.search(r"(사업자번호|대표자|주소|조회)", name):
            return name
    return ""


def extract_bizno_business_name(page_html: str, business_number: str) -> str:
    """Read the single BizNo search-result heading for the queried number."""
    for found_number, raw_name in re.findall(
        r'<a\s+href="/article/(\d+)"[^>]*>\s*<h4[^>]*>(.*?)</h4>\s*</a>',
        page_html,
        flags=re.IGNORECASE | re.DOTALL,
    ):
        if found_number != business_number:
            continue
        name = unescape(re.sub(r"<[^>]+>", "", raw_name)).strip()
        if 2 <= len(name) <= 80:
            return name
    return ""


class BizNoLookupClient:
    """Maintain a non-visible BizNo web session for business-name lookup.

    BizNo accepts the business number in its normal public ``query`` parameter.
    The initial page can be slow, so this client starts that request in a daemon
    thread at program launch.  The UI never waits for the warm-up unless the
    user initiates a lookup before it has completed.
    """

    def __init__(self) -> None:
        self._cookies = CookieJar()
        self._opener = build_opener(HTTPCookieProcessor(self._cookies))
        self._request_lock = Lock()
        self._warm_started = False
        self._warm_finished = Event()
        self._warm_error = ""

    def prewarm(self) -> None:
        if self._warm_started:
            return
        self._warm_started = True
        Thread(target=self._warm_up, name="bizno-prewarm", daemon=True).start()

    def _warm_up(self) -> None:
        try:
            self._request(BIZNO_URL, timeout_seconds=45)
        except Exception as exc:
            # A warm-up failure is not a user-facing failure. The first lookup
            # retries normally and reports its own concrete result.
            self._warm_error = str(exc)
        finally:
            self._warm_finished.set()

    def _request(self, url: str, timeout_seconds: int) -> str:
        request = Request(
            url,
            headers={
                "User-Agent": "EAccAutomation/1.0 (internal business lookup)",
                "Accept-Language": "ko-KR,ko;q=0.9",
            },
        )
        with self._request_lock:
            with self._opener.open(request, timeout=timeout_seconds) as response:
                if response.status != 200:
                    raise MerchantLookupError(f"비즈노 조회 응답이 HTTP {response.status}입니다.")
                return response.read().decode("utf-8", errors="replace")

    def lookup(self, business_number: str) -> str:
        if not is_valid_business_number(business_number):
            raise MerchantLookupError("사업자번호 형식 또는 체크섬이 올바르지 않습니다.")
        try:
            self.prewarm()
            # The first user click can arrive while prewarm is still fetching.
            # It is safe to wait because both requests share the same cookie jar.
            self._warm_finished.wait(timeout=45)
            url = f"{BIZNO_URL}?{urlencode({'query': business_number})}"
            page_html = self._request(url, timeout_seconds=45)
            name = extract_bizno_business_name(page_html, business_number)
            if not name:
                raise MerchantLookupError("비즈노 조회 결과에서 상호명을 하나로 확인하지 못했습니다.")
            return name
        except MerchantLookupError:
            raise
        except Exception as exc:
            raise MerchantLookupError(f"비즈노 조회에 실패했습니다: {exc}") from exc
