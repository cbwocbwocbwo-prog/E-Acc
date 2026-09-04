from __future__ import annotations

import asyncio
from io import BytesIO
import re
from pathlib import Path

from PIL import Image, ImageEnhance, ImageOps

from .models import (
    ReceiptFieldCheck,
    ReceiptValidationResult,
    UnsubmittedTransaction,
)


class OcrUnavailableError(RuntimeError):
    """Raised when the local Korean Windows OCR engine cannot be used."""


_OCR_ENGINE = None


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


def recognize_focused_bottom(paths: tuple[Path, ...], angle: int = 0) -> str:
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            text = await _recognize_image_bytes_async(_focused_bottom_bytes(path, angle))
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


def recognize_approval_lines(paths: tuple[Path, ...], angle: int = 0) -> str:
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            for payload in _approval_line_bytes(path, angle):
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
    return exact


def _date_check(transaction: UnsubmittedTransaction, text: str, tokens: tuple[str, ...]) -> ReceiptFieldCheck:
    year, month, day = transaction.evidence_date.split("-")
    # 2026년 08월20일, 2026-08-20, 2026기820처럼 구분자가 깨진 형식도 허용한다.
    # (,)처럼 흐린 0이 기호 둘로 분리되는 OCR 사례도 날짜 문맥에서만 보정한다.
    date_text = text.translate(
        str.maketrans({"O": "0", "o": "0", "Z": "2", "z": "2", "b": "6", "f": "6", "F": "6"})
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
        if re.search(r"POS|결제|일시|판매", context, flags=re.IGNORECASE):
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
        status="정상" if all(check.is_match for check in checks) else "이상",
        checks=checks,
        ocr_text=text,
    )


def recognize_images_until(
    transaction: UnsubmittedTransaction,
    paths: tuple[Path, ...],
    angle: int,
    base_text: str = "",
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
        for path in paths:
            for payload in _enhanced_region_bytes(path, angle, ((0.05, 0.48), (0.45, 1.0))):
                text = await _recognize_image_bytes_async(payload)
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
    
    # 실측 표본에서 가로로 긴 파일은 전부 270° 회전 저장본, 세로로 긴 파일은
    # 전부 정방향이었다. 파일 크기만 읽어 첫 판독 각도를 정한다(OCR 비용 0).
    first_angle = 0
    if image_paths:
        try:
            with Image.open(image_paths[0]) as probe:
                if probe.width > probe.height:
                    first_angle = 270
        except (OSError, ValueError):
            first_angle = 0
    second_angle = 270 if first_angle == 0 else 0

   # 승인번호가 주로 하단 카드결제 영역에 있는 사진은 이 단독 판독이 가장 안정적이다.
    focused_bottom_text = recognize_focused_bottom(image_paths, 0)
    upright_text = recognize_images(image_paths, 0)
    if upright_text.strip():
        upright = evaluate_ocr_text(transaction, upright_text)
        candidates.append((0, upright))
        # 이미 세 필드가 모두 일치하면 불필요한 세 번의 OCR을 생략한다.
        if upright.status == "정상":
            return upright
    if focused_bottom_text.strip():
        focused = evaluate_ocr_text(
            transaction,
            "\n".join(filter(None, (upright_text, focused_bottom_text))),
        )
        candidates.append((0, focused))
        if focused.status == "정상":
            return focused

    # 사진형 영수증의 카드결제 구간은 좁은 줄 단위로 한 번 더 읽는다. 이 보조 판독은
    # 승인번호를 명확히 읽을 수 있을 때만 이후의 정확 일치 검증을 통과한다.
    approval_line_text = recognize_approval_lines(image_paths, 0)
    if approval_line_text.strip():
        approval_lines = evaluate_ocr_text(
            transaction,
            "\n".join(filter(None, (upright_text, focused_bottom_text, approval_line_text))),
        )
        candidates.append((0, approval_lines))
        if approval_lines.status == "정상":
            return approval_lines

    # 실측 표본에서 회전 저장본은 전부 270°였다. 180°는 3단계 최후 시도로 미룬다.
    for angle in (270, 90):
        text = recognize_images(image_paths, angle)
        if not text.strip():
            continue
        rotated = evaluate_ocr_text(transaction, text)
        candidates.append((angle, rotated))
        # 1단계 upright/focused와 같은 조기 반환 규칙. 여기서 끝나면 3단계 27회×n을 아예 건너뛴다.
        if rotated.status == "정상":
            return ReceiptValidationResult(
                transaction_id=rotated.transaction_id,
                status=rotated.status,
                checks=rotated.checks,
                ocr_text=rotated.ocr_text,
                rotation_degrees=angle,
                orientation_ambiguous=False,
                orientation_reason="",
            )
    if not candidates:
        raise OcrUnavailableError("영수증에서 텍스트를 읽지 못했습니다.")


    # 사진은 정방향인데 회전 OCR이 우연히 숫자를 많이 읽어 잘못 선택되는 사례가 있다.
    # 반대로 세로·거꾸로 저장된 영수증은 기본 OCR의 점수가 낮아 강화 판독 후보에서
    # 빠지는 경우도 있다. 이상 건에 한해 네 방향 모두를 강화 판독해 필드 일치 수로
    # 최종 선택한다. 원본 파일은 전혀 수정하지 않는다.
    base_by_angle = {angle: result for angle, result in candidates}
    enhancement_angles = (270, 0, 90)
    enhanced_candidates: list[tuple[int, ReceiptValidationResult]] = []
    for angle in enhancement_angles:
        base = base_by_angle.get(angle)
        try:
            merged_text, merged = recognize_images_until(
                transaction,
                image_paths,
                angle,
                base.ocr_text if base else "",
            )
        except (OSError, ValueError):
            continue
        if not merged_text.strip():
            continue
        enhanced_candidates.append((angle, merged))
        # 강화 판독은 한 방향이 최대 OCR 27회다. 정상 판정이 나온 뒤에도 남은
        # 방향을 계속 돌면 정방향 영수증에서 81회가 순수 낭비된다. 앞의
        # upright·focused·approval_lines 단계와 같은 조기 반환 규칙을 적용한다.
        if merged.status == "정상":
            return ReceiptValidationResult(
                transaction_id=merged.transaction_id,
                status=merged.status,
                checks=merged.checks,
                ocr_text=merged.ocr_text,
                rotation_degrees=angle,
                orientation_ambiguous=False,
                orientation_reason="",
            )

    # 여기 도달했다는 건 세 방향 모두 정상이 아니라는 뜻이다. 거꾸로 저장된
    # 영수증만 남으므로 강화 판독 27회 대신 기본 OCR 1회로 안전망을 남긴다.
    try:
        flipped_text = recognize_images(image_paths, 180)
    except (OSError, ValueError):
        flipped_text = ""
    if flipped_text.strip():
        flipped = evaluate_ocr_text(transaction, flipped_text)
        candidates.append((180, flipped))
        if flipped.status == "정상":
            return ReceiptValidationResult(
                transaction_id=flipped.transaction_id,
                status=flipped.status,
                checks=flipped.checks,
                ocr_text=flipped.ocr_text,
                rotation_degrees=180,
                orientation_ambiguous=False,
                orientation_reason="",
            )



    final_candidates = [*candidates, *enhanced_candidates]
    ranked = sorted(final_candidates, key=lambda item: _orientation_score(item[1]), reverse=True)
    best_angle, best = ranked[0]
    best_score = _orientation_score(best)
    tied_angles = [angle for angle, result in ranked if _orientation_score(result) == best_score]
    ambiguous = len(tied_angles) > 1
    reason = ""
    if ambiguous:
        reason = "회전 방향 판정 점수가 같음: " + ", ".join(f"{angle}°" for angle in tied_angles)
    return ReceiptValidationResult(
        transaction_id=best.transaction_id,
        status="이상" if ambiguous else best.status,
        checks=best.checks,
        ocr_text=best.ocr_text,
        rotation_degrees=best_angle,
        orientation_ambiguous=ambiguous,
        orientation_reason=reason,
    )
