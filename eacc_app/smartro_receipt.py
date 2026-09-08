from __future__ import annotations

"""Supplemental OCR for the SmartroPAY receipt layout.

The general receipt verifier intentionally remains unchanged.  SmartroPAY
places its small approval number in the centre of a tall web receipt, so this
module reads only that labelled area at a much larger scale and merges a
confirmed approval number back into the normal validation result.
"""

import asyncio
from dataclasses import replace
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageEnhance, ImageOps

from . import ocr_validation
from .models import ReceiptFieldCheck, ReceiptValidationResult, UnsubmittedTransaction


_SMARTRO_MARKERS = ("smartropay", "smartro", "스마트로")


def _is_smartro_receipt(
    transaction: UnsubmittedTransaction,
    validation: ReceiptValidationResult,
) -> bool:
    source = " ".join((transaction.merchant, validation.ocr_text)).casefold()
    return any(marker in source for marker in _SMARTRO_MARKERS)


def _approval_check(result: ReceiptValidationResult) -> ReceiptFieldCheck | None:
    return next((check for check in result.checks if check.field_name == "승인번호"), None)


def _smartro_approval_payloads(path: Path, angle: int) -> tuple[bytes, ...]:
    """Return enlarged central SmartroPAY approval-number strips.

    SmartroPAY web receipts are vertically long.  The approval label/value lies
    around 31–42% of the height; the strips overlap so small browser zoom or
    downloaded-image padding does not move the number outside the crop.
    """
    with Image.open(path) as source:
        image = source.rotate(angle, expand=True) if angle else source.copy()
        grayscale = image.convert("L")
        width, height = grayscale.size
        variants: list[Image.Image] = []
        for start, end in ((0.27, 0.39), (0.31, 0.43), (0.25, 0.46)):
            strip = grayscale.crop(
                (int(width * 0.07), int(height * start), int(width * 0.95), int(height * end))
            )
            contrast = ImageEnhance.Contrast(ImageOps.autocontrast(strip, cutoff=1)).enhance(2.8)
            variants.extend((strip, contrast))
            # Keep a restrained threshold variant for faint grey SmartroPAY text.
            variants.append(contrast.point(lambda value: 255 if value > 185 else 0))

        payloads: list[bytes] = []
        for variant in variants:
            scale = min(7.0, 3600 / max(variant.width, variant.height))
            enlarged = variant.resize(
                (max(1, round(variant.width * scale)), max(1, round(variant.height * scale))),
                Image.Resampling.LANCZOS,
            )
            output = BytesIO()
            enlarged.save(output, format="PNG")
            payloads.append(output.getvalue())
        return tuple(payloads)


def _recognize_smartro_approval_text(paths: tuple[Path, ...], angle: int) -> str:
    async def recognize_all() -> str:
        texts: list[str] = []
        for path in paths:
            for payload in _smartro_approval_payloads(path, angle):
                text = await ocr_validation._recognize_image_bytes_async(payload)
                if text:
                    texts.append(text)
        return "\n".join(texts)

    return asyncio.run(recognize_all())


def merge_smartro_approval_result(
    transaction: UnsubmittedTransaction,
    validation: ReceiptValidationResult,
    focused_text: str,
) -> ReceiptValidationResult:
    """Merge a verified focused OCR approval number into *validation*.

    Only an exact match accepted by the existing verifier is used.  A failed
    focused read therefore cannot turn an abnormal receipt into a normal one.
    """
    if not focused_text.strip():
        return validation
    merged_text = "\n".join(filter(None, (validation.ocr_text, focused_text)))
    focused = ocr_validation.evaluate_ocr_text(transaction, merged_text)
    focused_approval = _approval_check(focused)
    existing_approval = _approval_check(validation)
    if (
        focused_approval is None
        or not focused_approval.is_match
        or existing_approval is None
        or existing_approval.is_match
    ):
        return validation

    confirmed = ReceiptFieldCheck(
        field_name=focused_approval.field_name,
        expected_value=focused_approval.expected_value,
        detected_value=focused_approval.detected_value,
        is_match=True,
        reason="SmartroPAY 승인번호 라벨 주변 보조 OCR 일치",
    )
    checks = tuple(
        confirmed if check.field_name == "승인번호" else check
        for check in validation.checks
    )
    return replace(
        validation,
        status="정상" if all(check.is_match for check in checks) else "이상",
        checks=checks,
        ocr_text=merged_text,
    )


def validate_receipt_with_smartro_support(
    transaction: UnsubmittedTransaction,
    image_paths: tuple[Path, ...],
) -> ReceiptValidationResult:
    """Run the unchanged general verifier, then SmartroPAY-focused OCR if needed."""
    validation = ocr_validation.validate_receipt_images(transaction, image_paths)
    approval = _approval_check(validation)
    if approval is None or approval.is_match or not _is_smartro_receipt(transaction, validation):
        return validation
    focused_text = _recognize_smartro_approval_text(image_paths, validation.rotation_degrees)
    return merge_smartro_approval_result(transaction, validation, focused_text)
