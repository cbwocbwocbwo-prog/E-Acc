"""Outlook 계정/스토어 진단 도구.

이 스크립트는 현재 이 PC의 Outlook 프로필에 등록된 모든 이메일 계정과,
각 계정의 배송 스토어(DeliveryStore)를 나열해서
"어떤 계정이 실제로 기본으로 사용되는지"를 명확히 보여줍니다.

메일이 왜 엉뚱한 계정으로 나가는지 진단할 때 사용하세요.

실행 방법 (PowerShell 또는 CMD):
    py -3.12 tools\\diagnose_outlook_accounts.py

또는 프로젝트 루트에서:
    py -3.12 -m tools.diagnose_outlook_accounts
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path


def _print_header(title: str) -> None:
    print()
    print("=" * 70)
    print(f" {title}")
    print("=" * 70)


def _read_mail_sender_json() -> dict | None:
    candidates = [
        Path.home() / "AppData" / "Local" / "EAccAutomation" / "mail_sender.json",
        Path(os.environ.get("LOCALAPPDATA", str(Path.home())))
        / "EAccAutomation"
        / "mail_sender.json",
    ]
    seen: set[str] = set()
    for path in candidates:
        key = str(path).casefold()
        if key in seen:
            continue
        seen.add(key)
        if not path.exists():
            print(f"  [없음] {path}")
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            print(f"  [있음] {path}")
            print(f"         내용: {data}")
            return data
        except Exception as exc:  # noqa: BLE001
            print(f"  [오류] {path} -> {exc}")
    return None


def main() -> int:
    _print_header("1. mail_sender.json 설정 파일 확인")
    settings = _read_mail_sender_json()
    requested = ""
    if settings:
        requested = str(settings.get("outlook_sender_address", "")).strip()
    print()
    if requested:
        print(f"  >> 요청된 발신 계정: {requested}")
    else:
        print("  >> 요청된 발신 계정: (지정 없음 -> 기본 계정 사용)")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print()
        print("[치명적] pywin32(win32com)가 설치되어 있지 않습니다.")
        print("  설치: py -3.12 -m pip install pywin32")
        return 1

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:  # noqa: BLE001
        print()
        print(f"[치명적] Outlook에 연결할 수 없습니다: {exc}")
        return 1

    namespace = outlook.GetNamespace("MAPI")

    _print_header("2. Outlook 프로필의 기본 스토어(DefaultStore)")
    try:
        default_store = namespace.DefaultStore
        print(f"  기본 스토어 이름: {default_store.DisplayName}")
        print(f"  StoreID (앞 60자): {str(default_store.StoreID)[:60]}...")
    except Exception as exc:  # noqa: BLE001
        print(f"  [오류] {exc}")

    _print_header("3. 등록된 모든 이메일 계정 (Session.Accounts)")
    accounts_info: list[dict] = []
    try:
        count = outlook.Session.Accounts.Count
    except Exception as exc:  # noqa: BLE001
        print(f"  [오류] 계정 목록을 읽을 수 없습니다: {exc}")
        return 1

    print(f"  총 {count}개 계정 등록됨")
    print()
    for index in range(1, count + 1):
        try:
            account = outlook.Session.Accounts.Item(index)
            smtp = str(account.SmtpAddress)
            display = str(account.DisplayName)
            try:
                store_name = str(account.DeliveryStore.DisplayName)
                store_id = str(account.DeliveryStore.StoreID)
            except Exception:
                store_name = "(없음)"
                store_id = ""
            accounts_info.append(
                {
                    "index": index,
                    "smtp": smtp,
                    "display": display,
                    "store_name": store_name,
                    "store_id": store_id,
                }
            )
            print(f"  [{index}] SMTP: {smtp}")
            print(f"      표시이름: {display}")
            print(f"      DeliveryStore: {store_name}")
            print()
        except Exception as exc:  # noqa: BLE001
            print(f"  [{index}] 오류: {exc}")

    _print_header("4. 진단 요약")
    default_store_id = ""
    try:
        default_store_id = str(namespace.DefaultStore.StoreID)
    except Exception:
        pass

    default_account = None
    for info in accounts_info:
        if info["store_id"] and info["store_id"] == default_store_id:
            default_account = info
            break

    if default_account:
        print(f"  >> 프로필 기본 계정 = {default_account['smtp']}")
    else:
        print("  >> 프로필 기본 계정을 특정할 수 없습니다.")

    if requested:
        match = next(
            (info for info in accounts_info if info["smtp"].casefold() == requested.casefold()),
            None,
        )
        if match is None:
            print()
            print(f"  [경고] 요청된 발신 계정 '{requested}'을(를) Outlook에서 찾을 수 없습니다.")
            print("         Outlook -> 파일 -> 계정 설정 에서 해당 계정을 추가해야 합니다.")
        else:
            print(f"  >> 요청된 발신 계정({requested})은 Outlook에 존재합니다. [OK]")
            if default_account and default_account["smtp"].casefold() != requested.casefold():
                print()
                print("  [원인 후보] 프로필 기본 계정과 요청 계정이 다릅니다.")
                print(f"     - 기본 계정: {default_account['smtp']}")
                print(f"     - 요청 계정: {requested}")
                print()
                print("     이 경우 Outlook은 SendUsingAccount를 무시하고")
                print("     기본 계정으로 발송하는 알려진 버그가 있습니다.")
                print()
                print("     [해결책 A - 권장] Outlook 기본 계정을 요청 계정으로 변경:")
                print("       Outlook -> 파일 -> 계정 설정 -> 계정 설정")
                print(f"       -> '{requested}' 선택 -> '기본값으로 설정' 클릭")
                print()
                print("     [해결책 B] 코드에서 요청 계정의 스토어에 직접 초안 생성")
                print("       (이번 커밋의 새 코드가 이 방식을 시도합니다)")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
