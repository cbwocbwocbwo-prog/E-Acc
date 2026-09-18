from __future__ import annotations

import asyncio
import atexit
from dataclasses import replace
from io import BytesIO
import json
from queue import Empty, Queue
import re
from pathlib import Path
import subprocess
import sys
from threading import Thread
import time

from PIL import Image, ImageEnhance, ImageFilter, ImageOps

from .models import (
    ReceiptFieldCheck,
    ReceiptValidationResult,
    UnsubmittedTransaction,
)


class OcrUnavailableError(RuntimeError):
    """Raised when the local Korean Windows OCR engine cannot be used."""


_OCR_ENGINE = None
_MONITOR_PHOTO_WORKER: subprocess.Popen[str] | None = None

# 일반 영수증의 빠른 기본 OCR은 제한하지 않는다. 이미 기본 OCR에서 정상으로
# 판정되지 않은 건만 아래의 보조 재판독 예산을 사용한다. Windows OCR 호출은
# 실행 중인 한 건을 강제로 취소할 수 없으므로, 호출 사이에서만 시간을 확인한다.
_FAILURE_RETRY_BUDGET_SECONDS = 15.0


def _deadline_reached(deadline: float | None) -> bool:
    return deadline is not None and time.monotonic() >= deadline


async def _recognize_image_bytes_async(data: bytes) -> str:
    global _OCR_ENGINE
    try:
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import BitmapDecoder
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    except ImportError as exc:
        raise OcrUnavailableError("Windows OCR 실행 모듈이 설치되지 않았습니다.") from exc

    # 엔진 생성은 언어 모델 초기화를 수반한다. 영수증 한 장에 OCR을 수십 번
    # 호출하므로 매번 만들면 그 비용이 그대로 누적된다. 같은 ko-KR 엔진을
    # 재사용하는 것이므로 인식 결과에는 영향이 없다.
    if _OCR_ENGINE is None:
        _OCR_ENGINE = OcrEngine.try_create_from_language(Language("ko-KR"))
    if _OCR_ENGINE is None:
        raise OcrUnavailableError("Windows 한국어 OCR 언어팩을 사용할 수 없습니다.")

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(data)
    await writer.store_async()
    writer.detach_stream()
    stream.seek(0)
    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    result = await _OCR_ENGINE.recognize_async(bitmap)
    return result.text or ""



def _rotated_image_bytes(path: Path, angle: int) -> bytes:
    if angle == 0:
        return path.read_bytes()
    with Image.open(path) as image:
        rotated = image.rotate(angle, expand=True)
        output = BytesIO()
        if rotated.mode not in {"RGB", "L"}:
            rotated = rotated.convert("RGB")
        rotated.save(output, format="JPEG", quality=95)
        return output.getvalue()


def _enhanced_region_bytes(path: Path, angle: int, region: tuple[float, float]) -> tuple[bytes, ...]:
    """Return contrast-enhanced whole/top/bottom OCR variants without changing the saved original."""
    with Image.open(path) as source:
        image = source.rotate(angle, expand=True) if angle else source.copy()
        grayscale = image.convert("L")
        width, height = grayscale.size
        variants: list[Image.Image] = []

        # 전체 강화본은 흐린 날짜·금액을, 상·하단 강화본은 카드전표의 승인번호를 보완한다.
        # 카드결제 전표는 하단 절반을 단독으로 읽을 때 승인번호가 가장 선명한
        # 경우가 있다. 이 두 가지는 이진화하지 않아 숫자 획을 보존한다.
        bottom = grayscale.crop((0, int(height * 0.50), width, height))
        variants.append(bottom)
        variants.append(
            ImageEnhance.Contrast(ImageOps.autocontrast(bottom, cutoff=1)).enhance(2.4)
        )
        full = ImageOps.autocontrast(grayscale, cutoff=1)
        variants.append(ImageEnhance.Contrast(full).enhance(2.0))

        # 웹 전표 캡처와 긴 카드 영수증은 승인번호·거래일시·합계가 각각 한 줄에
        # 배치된다. 큰 영역을 한꺼번에 읽으면 숫자열이 주변 글자와 합쳐지므로,
        # 겹치는 가로띠를 원본/대비 강화본으로 별도 판독한다.
        for start in (0.10, 0.24, 0.38, 0.52, 0.66):
            strip = grayscale.crop((0, int(height * start), width, int(height * min(start + 0.22, 1.0))))
            variants.append(strip)
            variants.append(
                ImageEnhance.Contrast(ImageOps.autocontrast(strip, cutoff=1)).enhance(2.6)
            )
        for start, end in region:
            cropped = grayscale.crop((0, int(height * start), width, int(height * end)))
            # 사진 영수증의 하단 카드결제 영역은 이진화하면 작은 숫자(승인번호·금액)가
            # 사라질 수 있다. 확대된 회색 원본과 대비 강화본도 함께 판독한다.
            variants.append(cropped)
            contrast_cropped = ImageEnhance.Contrast(
                ImageOps.autocontrast(cropped, cutoff=1)
            ).enhance(2.4)
            variants.append(contrast_cropped)
            # 조명·감열지 농도가 다양하므로 밝은 용지까지 포함해 다섯 단계 이진화로
            # 카드전표의 작은 승인번호·금액을 재판독한다.
            for threshold in (145, 165, 185, 205, 225):
                variants.append(
                    contrast_cropped.point(
                        lambda value, cutoff=threshold: 255 if value > cutoff else 0
                    )
                )

        payloads: list[bytes] = []
        for variant in variants:
            # Windows OCR은 지나치게 큰 비트맵을 거부한다. 확대하되 긴 변은 2,048px로 제한한다.
            scale = min(3.0, 2048 / max(variant.width, variant.height))
            target_size = (
                max(1, round(variant.width * scale)),
                max(1, round(variant.height * scale)),
            )
            enlarged = variant.resize(target_size, Image.Resampling.LANCZOS)
            output = BytesIO()
            enlarged.save(output, format="PNG")
            payloads.append(output.getvalue())
        return tuple(payloads)


def _focused_bottom_bytes(path: Path, angle: int) -> bytes:
    """Return the lower payment area as a single OCR image.

    Keeping this as a separate OCR call avoids the Windows OCR engine losing
    fine approval-number strokes while processing a long series of variants.
    """
    with Image.open(path) as source:
        image = source.rotate(angle, expand=True) if angle else source.copy()
        grayscale = image.convert("L")
        bottom = grayscale.crop((0, int(grayscale.height * 0.50), grayscale.width, grayscale.height))
        scale = min(3.0, 2048 / max(bottom.width, bottom.height))
        enlarged = bottom.resize(
            (max(1, round(bottom.width * scale)), max(1, round(bottom.height * scale))),
            Image.Resampling.LANCZOS,
        )
        output = BytesIO()
        enlarged.save(output, format="PNG")
        return output.getvalue()


def recognize_focused_bottom(
    paths: tuple[Path, ...],
    angle: int = 0,
    *,
    deadline: float | None = None,
) -> str:
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            if _deadline_reached(deadline):
                return "\n".join(texts)
            text = await _recognize_image_bytes_async(_focused_bottom_bytes(path, angle))
            if text:
                texts.append(text)
        return "\n".join(texts)

    return asyncio.run(recognize_all())


def _screen_capture_bytes(path: Path, angle: int) -> tuple[bytes, ...]:
    """Prepare small digital receipts and photos of a monitor for OCR.

    A browser receipt captured from a monitor is commonly only a few hundred
    pixels wide.  The ordinary OCR input intentionally preserves that original
    image for speed, but its thin approval-number glyphs can disappear.  This
    fallback enlarges a lightly sharpened full document and a margin-trimmed
    copy.  It is only called after the ordinary receipt path did not pass, so it
    never adds work to normal rows.
    """
    with Image.open(path) as source:
        image = source.rotate(angle, expand=True) if angle else source.copy()
        grayscale = ImageOps.exif_transpose(image).convert("L")
        width, height = grayscale.size
        trimmed = grayscale.crop(
            (int(width * 0.025), int(height * 0.02), int(width * 0.975), int(height * 0.98))
        )
        variants = (
            grayscale,
            ImageEnhance.Contrast(
                ImageOps.autocontrast(grayscale, cutoff=1).filter(
                    ImageFilter.UnsharpMask(radius=1.2, percent=135, threshold=2)
                )
            ).enhance(1.45),
            ImageEnhance.Contrast(
                ImageOps.autocontrast(trimmed, cutoff=1).filter(ImageFilter.MedianFilter(3))
            ).enhance(1.8),
        )
        payloads: list[bytes] = []
        for variant in variants:
            scale = min(4.5, 2400 / max(variant.width, variant.height))
            enlarged = variant.resize(
                (max(1, round(variant.width * scale)), max(1, round(variant.height * scale))),
                Image.Resampling.LANCZOS,
            )
            output = BytesIO()
            enlarged.save(output, format="PNG")
            payloads.append(output.getvalue())
        return tuple(payloads)


def recognize_screen_capture(
    paths: tuple[Path, ...],
    angle: int,
    *,
    deadline: float | None = None,
) -> str:
    """OCR a screen-captured or monitor-photographed receipt within a budget."""
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            for payload in _screen_capture_bytes(path, angle):
                if _deadline_reached(deadline):
                    return "\n".join(texts)
                text = await _recognize_image_bytes_async(payload)
                if text:
                    texts.append(text)
        return "\n".join(texts)

    return asyncio.run(recognize_all())


def recognize_monitor_photo(paths: tuple[Path, ...]) -> str:
    """Use isolated local Korean OCR for a failed monitor-photo receipt only."""
    if not paths:
        return ""
    # Windows OCR(winrt)과 EasyOCR(torch)는 같은 프로세스에서 DLL 초기화 충돌이
    # 날 수 있다. 별도 프로세스로 실행하면 정상 건의 메모리·처리속도에는 전혀
    # 영향을 주지 않으며, 실패한 화면 촬영본만 보조 OCR을 사용한다.
    worker = (
        "import easyocr, json, sys\n"
        "reader = easyocr.Reader(['ko', 'en'], gpu=False, verbose=False)\n"
        "for request in sys.stdin:\n"
        "    paths = json.loads(request)\n"
        "    texts = []\n"
        "    for image_path in paths:\n"
        "        texts.extend(reader.readtext(image_path, detail=0, paragraph=False))\n"
        "    print(json.dumps(texts, ensure_ascii=False), flush=True)\n"
    )
    try:
        process = _monitor_photo_worker(worker)
        assert process.stdin is not None and process.stdout is not None
        process.stdin.write(json.dumps([str(path) for path in paths]) + "\n")
        process.stdin.flush()
        response = _read_worker_response(process, timeout_seconds=30)
        if not response:
            return ""
        values = json.loads(response)
        return "\n".join(value for value in values if isinstance(value, str))
    except (OSError, RuntimeError, ValueError, json.JSONDecodeError):
        _stop_monitor_photo_worker()
        return ""


def _monitor_photo_worker(worker: str) -> subprocess.Popen[str]:
    global _MONITOR_PHOTO_WORKER
    if _MONITOR_PHOTO_WORKER is not None and _MONITOR_PHOTO_WORKER.poll() is None:
        return _MONITOR_PHOTO_WORKER
    _MONITOR_PHOTO_WORKER = subprocess.Popen(
        [sys.executable, "-c", worker],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return _MONITOR_PHOTO_WORKER


def _read_worker_response(
    process: subprocess.Popen[str], *, timeout_seconds: float
) -> str:
    assert process.stdout is not None
    responses: Queue[str] = Queue(maxsize=1)
    Thread(target=lambda: responses.put(process.stdout.readline()), daemon=True).start()
    try:
        return responses.get(timeout=timeout_seconds)
    except Empty:
        _stop_monitor_photo_worker()
        return ""


def _stop_monitor_photo_worker() -> None:
    global _MONITOR_PHOTO_WORKER
    process, _MONITOR_PHOTO_WORKER = _MONITOR_PHOTO_WORKER, None
    if process is not None and process.poll() is None:
        process.terminate()
        try:
            process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            process.kill()


atexit.register(_stop_monitor_photo_worker)


def _payment_detail_bytes(path: Path, angle: int) -> tuple[bytes, ...]:
    """Create a few high-contrast versions of the card-payment detail block.

    This is deliberately a *fallback* OCR source.  It is used only after the
    normal full-receipt OCR has already found the expected amount but could not
    read the approval number or date.  In particular, it helps photographed
    paper receipts whose lower approval block is faded or creased.
    """
    with Image.open(path) as source:
        image = source.rotate(angle, expand=True) if angle else source.copy()
        grayscale = image.convert("L")
        width, height = grayscale.size
        # The approval number, approval date/time and paid amount generally
        # occupy this lower-middle card-payment block.  Keep enough of the
        # labels so that Windows OCR does not interpret a number as a product
        # price, while excluding the mostly blank receipt margins.
        block = grayscale.crop(
            (int(width * 0.05), int(height * 0.48), int(width * 0.96), int(height * 0.92))
        )
        contrasted = ImageEnhance.Contrast(
            ImageOps.autocontrast(block, cutoff=1)
        ).enhance(2.8)
        # 하나카드 등 종이 영수증의 승인번호 줄은 하단 전체를 읽을 때 접힌
        # 선·공백에 묻히기 쉽다. 카드번호부터 결제일시까지의 왼쪽 결제정보
        # 블록을 별도로 확대하면 숫자와 항목명이 함께 남는다.
        approval_block = grayscale.crop(
            (int(width * 0.05), int(height * 0.66), int(width * 0.75), int(height * 0.84))
        )
        approval_auto = ImageOps.autocontrast(approval_block, cutoff=1)
        approval_contrasted = ImageEnhance.Contrast(approval_auto).enhance(2.0)
        variants = [
            block,
            contrasted,
            approval_auto,
            approval_contrasted,
        ]
        payloads: list[bytes] = []
        for variant in variants:
            scale = min(3.0, 2400 / max(variant.width, variant.height))
            enlarged = variant.resize(
                (max(1, round(variant.width * scale)), max(1, round(variant.height * scale))),
                Image.Resampling.LANCZOS,
            )
            output = BytesIO()
            enlarged.save(output, format="PNG")
            payloads.append(output.getvalue())
        return tuple(payloads)


def recognize_payment_details(
    paths: tuple[Path, ...],
    angle: int,
    *,
    deadline: float | None = None,
) -> str:
    """Read the lower payment block in a known receipt orientation."""
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            for payload in _payment_detail_bytes(path, angle):
                if _deadline_reached(deadline):
                    return "\n".join(texts)
                text = await _recognize_image_bytes_async(payload)
                if text:
                    texts.append(text)
        return "\n".join(texts)

    return asyncio.run(recognize_all())


def _approval_line_bytes(path: Path, angle: int) -> tuple[bytes, ...]:
    """Create enlarged lower-card-section strips for small approval numbers.

    A photographed receipt often has a long, mostly blank upper area.  Reading
    the complete lower half can make Windows OCR discard the thin digits on the
    ``승인 번호`` line.  The overlapping narrow strips keep that row large
    enough to recognize; the source image is never altered.
    """
    with Image.open(path) as source:
        image = source.rotate(angle, expand=True) if angle else source.copy()
        grayscale = image.convert("L")
        width, height = grayscale.size
        payloads: list[bytes] = []
        for start, end in ((0.70, 0.77), (0.745, 0.80), (0.78, 0.85)):
            strip = grayscale.crop(
                # 영수증의 좌우 여백을 조금 더 제외해야 세로 배경 무늬가 카드 숫자와
                # 합쳐지지 않는다. 일반 사진 영수증의 인쇄 영역을 보존하는 범위다.
                (int(width * 0.098), int(height * start), int(width * 0.916), int(height * end))
            )
            # A narrow strip is intentionally enlarged more than a whole receipt.
            scale = min(5.0, 3600 / max(strip.width, strip.height))
            enlarged = strip.resize(
                (max(1, round(strip.width * scale)), max(1, round(strip.height * scale))),
                Image.Resampling.LANCZOS,
            )
            output = BytesIO()
            enlarged.save(output, format="PNG")
            payloads.append(output.getvalue())
        return tuple(payloads)


def recognize_approval_lines(
    paths: tuple[Path, ...],
    angle: int = 0,
    *,
    deadline: float | None = None,
) -> str:
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            for payload in _approval_line_bytes(path, angle):
                if _deadline_reached(deadline):
                    return "\n".join(texts)
                text = await _recognize_image_bytes_async(payload)
                if text:
                    texts.append(text)
        return "\n".join(texts)

    return asyncio.run(recognize_all())


def _critical_transaction_detail_bytes(path: Path, angle: int) -> tuple[bytes, ...]:
    """Create narrow OCR inputs for date, approval number, and paid amount.

    This is a last-resort path for photographed card slips.  A full-page OCR
    can merge the thin numeric characters with a patterned background even
    when the three fields are plainly visible to a person.  The three bands
    keep the transaction timestamp, approval block, and total block separate.
    """
    with Image.open(path) as source:
        image = source.rotate(angle, expand=True) if angle else source.copy()
        grayscale = image.convert("L")
        width, height = grayscale.size
        payloads: list[bytes] = []

        # 일부 증빙은 영수증 자체가 아니라 모니터의 결제 팝업을 촬영한 사진이다.
        # 이 경우 화면 격자무늬가 숫자 획과 겹치므로, 중앙의 영수증 창만 잘라
        # 약한 중앙값 필터를 적용한 한 장을 최우선 보조 입력으로 사용한다.
        # 최후 재판독 단계에서만 생성되므로 정상 건에는 비용이 없다.
        screen_receipt = grayscale.crop(
            (int(width * 0.10), int(height * 0.14), int(width * 0.99), int(height * 0.92))
        )
        screen_receipt = ImageOps.autocontrast(
            screen_receipt.filter(ImageFilter.MedianFilter(3)), cutoff=1
        )
        screen_receipt = screen_receipt.resize(
            (max(1, screen_receipt.width * 2), max(1, screen_receipt.height * 2)),
            Image.Resampling.LANCZOS,
        )
        screen_output = BytesIO()
        screen_receipt.save(screen_output, format="PNG")
        payloads.append(screen_output.getvalue())

        # The bands deliberately overlap: a receipt layout can put its date
        # high in the body, while another puts the approval block near bottom.
        for start, end in ((0.20, 0.57), (0.44, 0.76), (0.64, 0.92)):
            band = grayscale.crop(
                (int(width * 0.04), int(height * start), int(width * 0.97), int(height * end))
            )
            variants = (
                band,
                ImageEnhance.Contrast(ImageOps.autocontrast(band, cutoff=1)).enhance(2.2),
            )
            for variant in variants:
                scale = min(5.0, 3600 / max(variant.width, variant.height))
                enlarged = variant.resize(
                    (max(1, round(variant.width * scale)), max(1, round(variant.height * scale))),
                    Image.Resampling.LANCZOS,
                )
                output = BytesIO()
                enlarged.save(output, format="PNG")
                payloads.append(output.getvalue())
        return tuple(payloads)


def recognize_critical_transaction_details(
    paths: tuple[Path, ...],
    angle: int,
    *,
    deadline: float | None = None,
    max_variants: int | None = None,
) -> str:
    """Read card-slip number fields after every regular OCR route has failed.

    Callers must invoke this only for an ``이상`` result.  It must never add
    latency to normal or 정상(2) receipts.
    """
    async def recognize_all() -> str:
        texts: list[str] = []
        attempted = 0
        for path in paths:
            for payload in _critical_transaction_detail_bytes(path, angle):
                if _deadline_reached(deadline) or (
                    max_variants is not None and attempted >= max_variants
                ):
                    return "\n".join(texts)
                text = await _recognize_image_bytes_async(payload)
                attempted += 1
                if text:
                    texts.append(text)
        return "\n".join(texts)

    return asyncio.run(recognize_all())


def _pos_datetime_strip_bytes(path: Path, angle: int) -> tuple[bytes, ...]:
    """Return the compact POS timestamp line from a photographed receipt.

    This is intentionally a single, late-stage OCR input.  On some landscape
    KICC receipts the date is printed immediately before ``POS`` but a full
    receipt pass merges the thin digits into the background.  Keeping just
    this line at readable scale preserves the date delimiter and adds no work
    to normal receipts.
    """
    with Image.open(path) as source:
        image = source.rotate(angle, expand=True) if angle else source.copy()
        grayscale = image.convert("L")
        width, height = grayscale.size
        strip = grayscale.crop(
            (int(width * 0.135), int(height * 0.418), int(width * 0.46), int(height * 0.47))
        )
        enlarged = strip.resize(
            (max(1, strip.width * 5), max(1, strip.height * 5)),
            Image.Resampling.LANCZOS,
        )
        output = BytesIO()
        enlarged.save(output, format="PNG")
        return (output.getvalue(),)


def recognize_pos_datetime_strip(
    paths: tuple[Path, ...],
    angle: int,
    *,
    deadline: float | None = None,
) -> str:
    """Read the POS timestamp only after every ordinary OCR route failed."""
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            for payload in _pos_datetime_strip_bytes(path, angle):
                if _deadline_reached(deadline):
                    return "\n".join(texts)
                text = await _recognize_image_bytes_async(payload)
                if text:
                    texts.append(text)
        return "\n".join(texts)

    return asyncio.run(recognize_all())


def recognize_images(
    paths: tuple[Path, ...],
    angle: int = 0,
    enhanced: bool = False,
) -> str:
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            if enhanced:
                # 하단 카드결제 정보의 제목 줄까지 포함해야 승인번호가 숫자열로
                # 분리되지 않는다. 중간을 약간 겹쳐 잘린 사진도 보완한다.
                payloads = _enhanced_region_bytes(path, angle, ((0.05, 0.48), (0.45, 1.0)))
            else:
                payloads = (_rotated_image_bytes(path, angle),)
            for payload in payloads:
                text = await _recognize_image_bytes_async(payload)
                if text:
                    texts.append(text)
        return "\n".join(text for text in texts if text)

    return asyncio.run(recognize_all())


def _numeric_tokens(text: str) -> tuple[str, ...]:
    # 숫자가 공백으로 끊기거나(/, I, l 등이 숫자로 잘못 읽히는 경우 포함)도
    # 승인번호 후보로 다시 조합한다. 원문 숫자 후보도 함께 보존한다.
    raw_tokens = re.findall(r"(?<!\d)\d[\d,./:\-]{1,20}\d(?!\d)|(?<!\d)\d(?!\d)", text)
    joined_tokens = re.findall(
        r"(?<![0-9A-Za-z])(?:[0-9OIl|/]{1,}(?:[ ,.:/-]+[0-9OIl|/]{1,})+|[0-9OIl|/]{4,})(?![0-9A-Za-z])",
        text,
    )
    translated = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "/": "7"})
    values: list[str] = []
    for token in (*raw_tokens, *joined_tokens):
        digits = re.sub(r"\D", "", token)
        if digits:
            values.append(digits)
        fuzzy_digits = re.sub(r"[^0-9]", "", token.translate(translated))
        if fuzzy_digits and fuzzy_digits != digits:
            values.append(fuzzy_digits)
    return tuple(dict.fromkeys(values))


def _field_check(
    field_name: str,
    expected_value: str,
    accepted_values: set[str],
    tokens: tuple[str, ...],
    allow_prefix: bool = False,
) -> ReceiptFieldCheck:
    found = next(
        (
            token
            for token in tokens
            if token in accepted_values
            or (allow_prefix and any(token.startswith(value) for value in accepted_values))
        ),
        "",
    )
    if found:
        return ReceiptFieldCheck(
            field_name=field_name,
            expected_value=expected_value,
            detected_value=found,
            is_match=True,
            reason="일치",
        )
    unique_candidates = list(dict.fromkeys(token for token in tokens if len(token) >= 4))[:8]
    candidates = ", ".join(unique_candidates)
    reason = f"영수증 OCR에서 화면값 {expected_value}을(를) 찾지 못했습니다."
    if candidates:
        reason += f" 숫자 후보: {candidates}"
    return ReceiptFieldCheck(
        field_name=field_name,
        expected_value=expected_value,
        detected_value=candidates,
        is_match=False,
        reason=reason,
    )


def _one_digit_edit_away(expected: str, candidate: str) -> bool:
    """Return true only for a single digit insertion, deletion, or replacement."""
    if abs(len(expected) - len(candidate)) > 1:
        return False
    if len(expected) == len(candidate):
        return sum(left != right for left, right in zip(expected, candidate)) == 1
    if len(expected) > len(candidate):
        expected, candidate = candidate, expected
    index = 0
    mismatches = 0
    while index < len(expected) and index < len(candidate):
        if expected[index] == candidate[index]:
            index += 1
            continue
        mismatches += 1
        if mismatches > 1:
            return False
        candidate = candidate[:index] + candidate[index + 1 :]
    return True


def _has_short_year_date_signature(transaction: UnsubmittedTransaction, text: str) -> bool:
    """Recognize a complete date when only the leading ``20`` was lost by OCR.

    This is deliberately *not* a general evidence-date match.  It is used
    solely as an additional guard for the narrow card-payment approval-number
    recovery below: exact month/day and the final two year digits must still
    be present.  A bare ``26-09-14`` is therefore never enough to make a
    receipt pass by itself.
    """
    year, month, day = transaction.evidence_date.split("-")
    repaired = text.translate(str.maketrans({"니": "1", "이": "9"}))
    year_pattern = rf"(?:{re.escape(year)}|{re.escape(year[-2:])})"
    pattern = (
        rf"(?<!\d){year_pattern}\D{{0,6}}0?{int(month)}"
        rf"\D{{0,6}}0?{int(day)}(?!\d)"
    )
    return bool(re.search(pattern, repaired))


def _approval_check(transaction: UnsubmittedTransaction, text: str, tokens: tuple[str, ...]) -> ReceiptFieldCheck:
    expected = transaction.approval_number
    exact = _field_check("승인번호", expected, {expected}, tokens)
    if exact.is_match:
        return exact

    # 주유소 감열지의 얇은 숫자는 Windows OCR에서 한자/영문자로 바뀌는 경우가 있다.
    # 예: 22502602 -> 牙25d2名02. 일반 본문에 이 치환을 적용하면 오검증 위험이 있으므로,
    # 반드시 '승인 번호' 라벨 바로 뒤의 짧은 값만 대상으로 하고 최종 값이 화면 승인번호와
    # 정확히 같을 때만 인정한다.
    approval_label_candidates = re.findall(
        r"(?:승\s*인|인)\s*(?:번\s*호|번)\s*[:：]?\s*([^\s]{4,12})",
        text,
        flags=re.IGNORECASE,
    )
    approval_lookalikes = str.maketrans(
        {
            "O": "0",
            "o": "0",
            "I": "1",
            "l": "1",
            "|": "1",
            "/": "7",
            # 실제 주유 영수증 OCR 오인 사례. 승인번호 라벨 문맥으로 엄격히 한정한다.
            "牙": "2",
            "名": "6",
            "d": "0",
            "D": "0",
        }
    )
    for candidate in approval_label_candidates:
        repaired = candidate.translate(approval_lookalikes)
        repaired_digits = re.sub(r"\D", "", repaired)
        if repaired_digits == expected:
            return ReceiptFieldCheck(
                field_name="승인번호",
                expected_value=expected,
                detected_value=repaired_digits,
                is_match=True,
                reason="승인번호 라벨 문맥의 숫자 OCR 보정 일치",
            )

    # KICC 카드전표에는 ``[승인번호] KICC...`` 형식의 줄이 별도로 인쇄된다.
    # 이 거래의 실제 OCR처럼 ``[I箋04411] K표:로제출``로 읽힐 수 있다. 여기서
    # ``I``은 1, ``箋``은 붙어 인식된 71, ``D``는 0의 오인 사례다. 이 보정은
    # KICC 바로 앞의 대괄호 값에만 적용하고, 보정 뒤 전체 숫자열이 e-Acc
    # 승인번호와 정확히 같을 때만 통과시킨다. 뒷자리(예: 4411)만 같은 경우는
    # 절대로 통과하지 않는다.
    kicc_candidates = re.findall(
        # KICC의 마지막 C도 OCR에서 J 등 한 글자로 바뀔 수 있고, 실제 전표에서는
        # ``KICC로제출`` 전체가 ``K표:로제출``로 바뀌기도 한다. 어느 경우든
        # 대괄호 값 바로 뒤의 KICC/로제출 문맥을 모두 갖춰야 한다.
        r"\[([^\]\r\n]{4,16})\]\s*(?:KIC(?:C|[A-Za-z가-힣])|K[가-힣][^\r\n]{0,4}로\s*제출)",
        text,
        flags=re.IGNORECASE,
    )
    kicc_lookalikes = str.maketrans(
        {
            "O": "0",
            "o": "0",
            "I": "1",
            "l": "1",
            "|": "1",
            "/": "7",
            "D": "0",
            "d": "0",
            # KICC 줄의 실제 Windows OCR 오인 사례: 한 글자에 '71'이 합쳐진다.
            "箋": "71",
            # 같은 글꼴에서 드물게 보이는 대체 인식값도 KICC 문맥으로 한정한다.
            "笋": "71",
        }
    )
    for candidate in kicc_candidates:
        repaired = candidate.translate(kicc_lookalikes)
        repaired_digits = re.sub(r"\D", "", repaired)
        if repaired_digits == expected:
            return ReceiptFieldCheck(
                field_name="승인번호",
                expected_value=expected,
                detected_value=repaired_digits,
                is_match=True,
                reason="KICC 카드전표 승인번호 OCR 보정 일치",
            )

    # OCR은 대괄호·영문자와 붙은 8자리 승인번호의 한 글자를 자주 덧붙이거나 놓친다.
    # 일반 숫자 전체에 적용하면 날짜·금액을 오인할 수 있으므로, 승인 문맥에서만 허용한다.
    context_parts = re.findall(
        r".{0,36}(?:승인|결제|거래|KIC[,. ]?C|VAN).{0,36}",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    context_text = "\n".join(context_parts)
    context_tokens = list(_numeric_tokens(context_text))
    # 감열지의 90이 %로 오인되는 사례가 있다. 승인 문맥에서만 %를 90으로
    # 해석하고, 아래의 한 자리 편집 보정까지 통과해야 일치로 인정한다.
    for raw_token in re.findall(r"(?<![0-9A-Za-z])[0-9OIl|/%]{4,}(?![0-9A-Za-z])", context_text):
        repaired = raw_token.translate(
            str.maketrans(
                {"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "/": "7", "%": "90"}
            )
        )
        digits = re.sub(r"\D", "", repaired)
        if digits:
            context_tokens.append(digits)
    fuzzy = next(
        (
            token
            for token in context_tokens
            if _one_digit_edit_away(expected, token)
        ),
        "",
    )
    if fuzzy:
        return ReceiptFieldCheck(
            field_name="승인번호",
            expected_value=expected,
            detected_value=fuzzy,
            is_match=True,
            reason="승인 문맥의 한 자리 OCR 보정 일치",
        )

    # 사진형 카드전표의 하단은 숫자 한 글자가 틀어지는 경우가 있다. 이 보정은
    # 카드/결제 문맥 안에서 (1) 화면 금액이 정확히 같이 읽히고, (2) 해당 영수증의
    # 연도 끝 두 자리·월·일도 모두 남아 있을 때에만 허용한다. 즉 ``26-09-니4``
    # 자체를 날짜 일치로 완화하지 않으며, 잘못 첨부된 다른 영수증을 한 자리
    # 승인번호 오인만으로 통과시키지 않는다.
    expected_amount = str(int(transaction.amount))
    card_payment_contexts = re.findall(
        r".{0,72}(?:[카가]드|결제|KIC[,. ]?C|VAN).{0,120}",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if _has_short_year_date_signature(transaction, text):
        for context in card_payment_contexts:
            if expected_amount not in _numeric_tokens(context):
                continue
            for candidate in _numeric_tokens(context):
                if len(candidate) != len(expected):
                    continue
                if _one_digit_edit_away(expected, candidate):
                    return ReceiptFieldCheck(
                        field_name="승인번호",
                        expected_value=expected,
                        detected_value=candidate,
                        is_match=True,
                        reason="카드결제·금액·날짜 문맥의 한 자리 OCR 보정 일치",
                    )
    return exact


def _date_check(transaction: UnsubmittedTransaction, text: str, tokens: tuple[str, ...]) -> ReceiptFieldCheck:
    year, month, day = transaction.evidence_date.split("-")
    # 2026년 08월20일, 2026-08-20, 2026기820처럼 구분자가 깨진 형식도 허용한다.
    # (,)처럼 흐린 0이 기호 둘로 분리되는 OCR 사례도 날짜 문맥에서만 보정한다.
    date_text = text.translate(
        str.maketrans(
            {
                "O": "0", "o": "0", "Z": "2", "z": "2", "b": "6", "f": "6", "F": "6",
                # 모니터 촬영 영수증에서 9가 한글 '이'로 읽힌 사례:
                # 2026.09.09 -> 2026.0이09. 전체 날짜 패턴이 일치할 때만
                # 인정되므로 이 보정만으로 임의 날짜를 통과시키지는 않는다.
                "이": "9",
                # 감열지의 9가 '누'로 인식된 사례. 아래의 전체 날짜 또는
                # POS 문맥의 월·일 일치 조건을 모두 통과할 때만 사용한다.
                "누": "9",
            }
        )
    )
    date_text = date_text.replace("()", "0").replace(",)", "0").replace("(,", "0")
    pattern = rf"{year}\D{{0,5}}0?{int(month)}\D{{0,5}}0?{int(day)}"
    if re.search(pattern, date_text):
        return ReceiptFieldCheck(
            field_name="증빙일자",
            expected_value=transaction.evidence_date,
            detected_value=transaction.evidence_date.replace("-", ""),
            is_match=True,
            reason="일치",
        )

    # 연도 부분만 흐려진 사진 영수증도 있다. 월·일이 일치하고 해당 숫자가
    # 결제/승인 시각 문맥에 있을 때만 허용해, 품목번호 등을 날짜로 오인하지 않는다.
    month_day_pattern = rf"(?<!\d)0?{int(month)}\D{{0,3}}0?{int(day)}(?!\d)"
    for match in re.finditer(month_day_pattern, date_text):
        context = date_text[max(0, match.start() - 90) : match.end() + 90]
        # date_text에서는 숫자 보정 때문에 POS가 P0S로 바뀔 수 있다.
        # 따라서 이 결제 시각 문맥도 함께 인정한다.
        if re.search(r"P[O0]S|결제|일시|판매", context, flags=re.IGNORECASE):
            return ReceiptFieldCheck(
                field_name="증빙일자",
                expected_value=transaction.evidence_date,
                detected_value=f"{month:0>2}{day:0>2}",
                is_match=True,
                reason="결제 시각 문맥의 월·일 OCR 보정 일치",
            )
    date_digits = transaction.evidence_date.replace("-", "")
    return _field_check(
        "증빙일자",
        transaction.evidence_date,
        {date_digits, date_digits[2:]},
        tokens,
        allow_prefix=True,
    )


def _amount_check(transaction: UnsubmittedTransaction, tokens: tuple[str, ...]) -> ReceiptFieldCheck:
    expected = str(int(transaction.amount))
    direct = _field_check("사용금액", f"{transaction.amount:,.0f}", {expected}, tokens)
    if direct.is_match:
        return direct

    # 합계금액이 흐리게 인식된 경우에도 공급가액과 부가세가 정확히 읽히면 합계로 검증한다.
    amounts = sorted({int(token) for token in tokens if 0 < int(token) <= int(transaction.amount)})
    for index, first in enumerate(amounts):
        for second in amounts[index + 1 :]:
            if first + second == int(transaction.amount):
                return ReceiptFieldCheck(
                    field_name="사용금액",
                    expected_value=f"{transaction.amount:,.0f}",
                    detected_value=f"{first:,} + {second:,}",
                    is_match=True,
                    reason="공급가액과 부가세의 합계가 일치",
                )
    return direct


def receipt_validation_status(checks: tuple[ReceiptFieldCheck, ...]) -> str:
    """Classify the three mandatory receipt fields without hiding a two-field pass.

    The amount must always match.  A receipt is accepted as ``정상(2)`` only
    when that amount is accompanied by the matching approval number *or* the
    matching evidence date.  Approval number plus date alone never passes.
    """
    matched_fields = {check.field_name for check in checks if check.is_match}
    if {"승인번호", "증빙일자", "사용금액"} <= matched_fields:
        return "정상"
    if "사용금액" in matched_fields and (
        "승인번호" in matched_fields or "증빙일자" in matched_fields
    ):
        return "정상(2)"
    return "이상"


def evaluate_ocr_text(
    transaction: UnsubmittedTransaction,
    text: str,
) -> ReceiptValidationResult:
    tokens = _numeric_tokens(text)
    checks = (
        _approval_check(transaction, text, tokens),
        _date_check(transaction, text, tokens),
        _amount_check(transaction, tokens),
    )
    return ReceiptValidationResult(
        transaction_id=transaction.transaction_id,
        status=receipt_validation_status(checks),
        checks=checks,
        ocr_text=text,
    )


def recognize_images_until(
    transaction: UnsubmittedTransaction,
    paths: tuple[Path, ...],
    angle: int,
    base_text: str = "",
    *,
    deadline: float | None = None,
    max_variants: int | None = None,
) -> tuple[str, ReceiptValidationResult]:
    """Read enhanced variants one by one, stopping as soon as all fields match.

    Every field check searches for a candidate in the accumulated text, so the
    verdict is monotone: text added later can turn 이상 into 정상 but never the
    reverse.  Stopping early therefore yields the same verdict the full 27-
    variant pass would, and no preprocessing variant is removed — the remaining
    ones are simply not needed once the receipt is already confirmed.
    """
    texts: list[str] = [base_text] if base_text else []

    async def run() -> str:
        attempted = 0
        for path in paths:
            for payload in _enhanced_region_bytes(path, angle, ((0.05, 0.48), (0.45, 1.0))):
                if _deadline_reached(deadline) or (
                    max_variants is not None and attempted >= max_variants
                ):
                    return "\n".join(texts)
                text = await _recognize_image_bytes_async(payload)
                attempted += 1
                if not text:
                    continue
                texts.append(text)
                if evaluate_ocr_text(transaction, "\n".join(texts)).status == "정상":
                    return "\n".join(texts)
        return "\n".join(texts)

    merged_text = asyncio.run(run())
    return merged_text, evaluate_ocr_text(transaction, merged_text)




def _orientation_score(result: ReceiptValidationResult) -> tuple[int, int, int]:
    match_count = sum(check.is_match for check in result.checks)
    keyword_count = sum(
        result.ocr_text.count(keyword)
        for keyword in ("승인", "금액", "합계", "일자", "사업자", "카드", "상호")
    )
    readable_count = len(re.sub(r"\s", "", result.ocr_text))
    return match_count, keyword_count, readable_count


def validate_receipt_images(
    transaction: UnsubmittedTransaction,
    image_paths: tuple[Path, ...],
) -> ReceiptValidationResult:
    candidates: list[tuple[int, ReceiptValidationResult]] = []
    # 이 시각부터는 정상 판정에 실패한 한 건 전체가 공유하는 보조 OCR 시간
    # 예산이다. 정상 건은 첫 OCR에서 바로 반환하므로 이 제한의 영향을 받지 않는다.
    failure_retry_deadline = time.monotonic() + _FAILURE_RETRY_BUDGET_SECONDS

    def is_eligible(result: ReceiptValidationResult) -> bool:
        return result.status in {"정상", "정상(2)"}

    # 가장 빠른 원본 전체 OCR을 먼저 실행한다. 통과하면 아래의 사진 보조 OCR은
    # 전혀 실행하지 않는다.
    upright_text = recognize_images(image_paths, 0)
    if upright_text.strip():
        upright = replace(
            evaluate_ocr_text(transaction, upright_text),
            account_validation_text=upright_text,
        )
        candidates.append((0, upright))
        # 정상(2)도 자동 결재 가능한 판정이다. 더 많은 OCR을 실행해 세 번째
        # 필드를 찾으려다 정상 행의 속도를 늦추지 않는다.
        if is_eligible(upright):
            return upright

    # 모니터를 촬영한 승인전표는 화면 주사선·기울어짐 때문에 Windows OCR이
    # 숫자만 깨뜨릴 수 있다. 기본 OCR 실패 건에만 먼저 한국어 보조 OCR을 적용해
    # 일반 보조 OCR이 시간 예산을 모두 쓰기 전에 이 유형을 복구한다.
    monitor_photo_text = recognize_monitor_photo(image_paths)
    if monitor_photo_text.strip():
        monitor_photo = replace(
            evaluate_ocr_text(
                transaction,
                "\n".join(filter(None, (upright_text, monitor_photo_text))),
            ),
            account_validation_text=upright_text,
        )
        candidates.append((0, monitor_photo))
        if is_eligible(monitor_photo):
            return monitor_photo

    # 승인번호가 주로 하단 카드결제 영역에 있는 사진은 이 단독 판독이 가장 안정적이다.
    focused_bottom_text = recognize_focused_bottom(
        image_paths,
        0,
        deadline=failure_retry_deadline,
    )
    if focused_bottom_text.strip():
        focused = replace(
            evaluate_ocr_text(
                transaction,
                "\n".join(filter(None, (upright_text, focused_bottom_text))),
            ),
            account_validation_text=upright_text,
        )
        candidates.append((0, focused))
        if is_eligible(focused):
            return focused

    # 사진형 영수증의 카드결제 구간은 좁은 줄 단위로 한 번 더 읽는다. 이 보조 판독은
    # 승인번호를 명확히 읽을 수 있을 때만 이후의 정확 일치 검증을 통과한다.
    approval_line_text = recognize_approval_lines(
        image_paths,
        0,
        deadline=failure_retry_deadline,
    )
    if approval_line_text.strip():
        approval_lines = replace(
            evaluate_ocr_text(
                transaction,
                "\n".join(filter(None, (upright_text, focused_bottom_text, approval_line_text))),
            ),
            account_validation_text=upright_text,
        )
        candidates.append((0, approval_lines))
        if is_eligible(approval_lines):
            return approval_lines

    # 가로 저장 또는 옆으로 촬영된 영수증만 회전 기본 OCR을 시도한다.
    for angle in (270, 90):
        if _deadline_reached(failure_retry_deadline):
            break
        text = recognize_images(image_paths, angle)
        if not text.strip():
            continue
        rotated = replace(
            evaluate_ocr_text(transaction, text),
            account_validation_text=text,
        )
        candidates.append((angle, rotated))
        if is_eligible(rotated):
            return replace(
                rotated,
                rotation_degrees=angle,
                orientation_ambiguous=False,
                orientation_reason="",
            )

        # 정상 영수증은 위의 전체 OCR에서 바로 반환한다. 금액만 확인된 실패 건에
        # 한해서만 같은 회전 방향의 결제정보 영역을 네 가지 대비로 재판독한다.
        # 따라서 일반적인 정상 건에는 추가 OCR 호출이나 처리 지연이 없다.
        amount_matched = any(
            check.field_name == "사용금액" and check.is_match
            for check in rotated.checks
        )
        if not amount_matched:
            continue
        payment_detail_text = recognize_payment_details(
            image_paths,
            angle,
            deadline=failure_retry_deadline,
        )
        if not payment_detail_text.strip():
            continue
        payment_detail = replace(
            evaluate_ocr_text(
                transaction,
                "\n".join(filter(None, (text, payment_detail_text))),
            ),
            account_validation_text=text,
        )
        candidates.append((angle, payment_detail))
        if is_eligible(payment_detail):
            return replace(
                payment_detail,
                rotation_degrees=angle,
                orientation_ambiguous=False,
                orientation_reason="",
            )
    # 뒤집힌 사진은 한 번만 확인한다. 과거처럼 세 방향 각각 수십 장을 강화 OCR
    # 하지 않으므로, 판독할 수 없는 영수증 한 건이 다음 행을 오래 막지 않는다.
    try:
        flipped_text = (
            recognize_images(image_paths, 180)
            if not _deadline_reached(failure_retry_deadline)
            else ""
        )
    except (OSError, ValueError):
        flipped_text = ""
    if flipped_text.strip():
        flipped = replace(
            evaluate_ocr_text(transaction, flipped_text),
            account_validation_text=flipped_text,
        )
        candidates.append((180, flipped))
        if is_eligible(flipped):
            return replace(
                flipped,
                rotation_degrees=180,
                orientation_ambiguous=False,
                orientation_reason="",
            )

    if not candidates:
        raise OcrUnavailableError("영수증에서 텍스트를 읽지 못했습니다.")

    ranked = sorted(candidates, key=lambda item: _orientation_score(item[1]), reverse=True)
    best_angle, best = ranked[0]

    # 여기부터는 기본 OCR이 통과하지 못한 예외 후보에만 적용하는 최대 15초
    # 보조 경로다. 작은 웹 영수증·모니터 촬영본은 원본 해상도로 읽으면 승인번호
    # 획이 사라지므로, 먼저 전체 문서를 확대·선명화한 입력을 사용한다.
    screen_text = recognize_screen_capture(
        image_paths,
        best_angle,
        deadline=failure_retry_deadline,
    )
    if screen_text.strip():
        screen = replace(
            evaluate_ocr_text(
                transaction,
                "\n".join(filter(None, (best.ocr_text, screen_text))),
            ),
            account_validation_text=best.account_validation_text,
        )
        candidates.append((best_angle, screen))
        if is_eligible(screen):
            return replace(
                screen,
                rotation_degrees=best_angle,
                orientation_ambiguous=False,
                orientation_reason="",
            )
        best = screen

    # 화면 캡처 보정 뒤에도 사용금액 자체가 보이지 않으면, 마지막 숫자영역
    # 재판독은 세 장으로 제한한다. 금액이 보인 건만 전체 좁은 영역을 시도한다.
    amount_matched = any(
        check.field_name == "사용금액" and check.is_match for check in best.checks
    )
    if best.status == "이상":
        critical_text = recognize_critical_transaction_details(
            image_paths,
            best_angle,
            deadline=failure_retry_deadline,
            max_variants=None if amount_matched else 3,
        )
        if critical_text.strip():
            critical = evaluate_ocr_text(
                transaction,
                "\n".join(filter(None, (best.ocr_text, critical_text))),
            )
            if is_eligible(critical):
                return replace(
                    critical,
                    account_validation_text=best.account_validation_text,
                    rotation_degrees=best_angle,
                    orientation_ambiguous=False,
                    orientation_reason="",
                )

        # 마지막으로 POS 시각행 하나만 크게 읽는다. 앞 단계와 같은 제한 시간
        # 안에서만 실행되므로 실패 건의 전체 처리 시간은 통제된다.
        pos_datetime_text = ""
        if amount_matched and not _deadline_reached(failure_retry_deadline):
            pos_datetime_text = recognize_pos_datetime_strip(
                image_paths,
                best_angle,
                deadline=failure_retry_deadline,
            )
        if pos_datetime_text.strip():
            pos_datetime = evaluate_ocr_text(
                transaction,
                "\n".join(filter(None, (best.ocr_text, critical_text, pos_datetime_text))),
            )
            if is_eligible(pos_datetime):
                return replace(
                    pos_datetime,
                    account_validation_text=best.account_validation_text,
                    rotation_degrees=best_angle,
                    orientation_ambiguous=False,
                    orientation_reason="",
                )

    final_ranked = sorted(candidates, key=lambda item: _orientation_score(item[1]), reverse=True)
    best_angle, best = final_ranked[0]
    best_score = _orientation_score(best)
    tied_angles = [angle for angle, result in final_ranked if _orientation_score(result) == best_score]
    ambiguous = len(tied_angles) > 1
    reason = ""
    if ambiguous:
        reason = "회전 방향 판정 점수가 같음: " + ", ".join(f"{angle}°" for angle in tied_angles)
    return ReceiptValidationResult(
        transaction_id=best.transaction_id,
        status="이상" if ambiguous else best.status,
        checks=best.checks,
        ocr_text=best.ocr_text,
        account_validation_text=best.account_validation_text,
        rotation_degrees=best_angle,
        orientation_ambiguous=ambiguous,
        orientation_reason=reason,
    )
