"""One-time, interactive diagnosis for the current e-Acc first four rows.

This file never stores login credentials.  It is intentionally not part of the
application's normal launch path.
"""
from __future__ import annotations

from getpass import getpass
from pathlib import Path
from threading import Event

from eacc_app.account_validation import validate_account_rules
from eacc_app.browser_automation import EAccountingBrowserService, LoginCredentials
from eacc_app.employee_directory import load_employee_names
from eacc_app.merchant_lookup import is_pg_business_type
from eacc_app.ocr_validation import validate_receipt_images


def invoke(start):
    done = Event()
    box: dict[str, object] = {}

    def success(value):
        box["value"] = value
        done.set()

    def failure(error):
        box["error"] = error
        done.set()

    start(success, failure)
    if not done.wait(90):
        raise TimeoutError("e-Acc 자동화 응답이 90초 안에 오지 않았습니다.")
    if "error" in box:
        raise box["error"]  # type: ignore[misc]
    return box["value"]


def main() -> None:
    user_id = input("i-NET 아이디: ").strip()
    password = getpass("i-NET 비밀번호: ")
    credentials = LoginCredentials(user_id=user_id, password=password)
    service = EAccountingBrowserService(Path.home() / "AppData" / "Local" / "EAccAutomation" / "downloads")
    excluded: set[str] = set()
    targets = []
    try:
        for index in range(1, 5):
            transaction = invoke(
                lambda ok, err: service.read_current_first_target(
                    credentials,
                    ok,
                    err,
                    excluded_transaction_ids=frozenset(excluded),
                )
            )
            excluded.add(transaction.transaction_id)
            targets.append(transaction)
            print(
                f"[행 {index}] 승인={transaction.approval_number} / "
                f"업종={transaction.business_type} / 계정={transaction.account_name} / "
                f"증빙={transaction.evidence_status} / 금액={transaction.amount}"
            )

        try:
            employee_names = load_employee_names()
        except Exception as exc:
            employee_names = ()
            print(f"직원 명단을 불러오지 못함: {exc}")

        for index, transaction in enumerate(targets, start=1):
            if "PD일반" in transaction.business_type:
                print(f"[행 {index}] PD일반 제외")
                continue
            if transaction.evidence_status != "#":
                print(f"[행 {index}] 결재 미진행: 증빙유무={transaction.evidence_status}")
                continue
            receipt = invoke(
                lambda ok, err, tx=transaction: service.download_receipt(tx, credentials, ok, err)
            )
            validation = validate_receipt_images(transaction, receipt.image_paths)
            account = validate_account_rules(transaction, employee_names, validation.ocr_text)
            print(
                f"[행 {index}] OCR={validation.status} / "
                f"계정판정={account.status}{(': ' + account.reason_text) if account.reason_text else ''}"
            )
            if validation.status != "정상" or account.status != "정상":
                print(f"[행 {index}] 결재 미진행: 기존 검증 기준 미통과")
                continue
            if is_pg_business_type(transaction.business_type):
                print(f"[행 {index}] PG일반: 실구매처 등록 전이므로 결재 미진행")
                continue
            opened = invoke(
                lambda ok, err, tx=transaction: service.open_approval_line(tx, credentials, ok, err)
            )
            print(f"[행 {index}] 결재선: {opened}")
            submitted = invoke(
                lambda ok, err, tx=transaction: service.submit_prepared_approval(tx, ok, err)
            )
            print(f"[행 {index}] 결재요청: {submitted}")
    finally:
        # Credentials and the temporary Edge profile are discarded with service.close().
        password = ""
        service.close()


if __name__ == "__main__":
    main()
