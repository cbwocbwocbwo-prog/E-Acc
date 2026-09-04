from __future__ import annotations

import os
import threading
import tkinter as tk
from dataclasses import replace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from .account_validation import AccountValidationResult, validate_account_rules
from .browser_automation import (
    ActualMerchantRegistrationError,
    ApprovalInProgressError,
    BrowserAutomationError,
    BrowserLoginRequired,
    EAccountingBrowserService,
    LoginCredentials,
    NoEligibleTransactions,
    NoUnsubmittedTransactions,
    ReceiptNotAvailable,
)
from .employee_directory import EmployeeDirectoryError, load_employee_names
from .merchant_lookup import (
    BizNoLookupClient,
    MerchantLookupError,
    extract_business_numbers,
    is_pg_business_type,
)
from .models import (
    DISPLAY_HEADERS,
    ImportDisplayRow,
    ImportSummary,
    ReceiptImageResult,
    ReceiptValidationResult,
    MerchantLookupResult,
    ProcessingEvent,
    UnsubmittedTransaction,
)
from .ocr_validation import validate_receipt_images
from .parser import WorkbookFormatError
from .storage import ImportRepository


APP_TITLE = "E-Acc 법인카드 자동처리"
STATUS_FILTERS = ("전체", "신규", "중복", "오류")
# 거래내역의 '처리결과'는 영수증 OCR·PG 조회와 구분된 결재 처리 상태만 표시한다.
APPROVAL_RESULT_STATUSES = frozenset(
    {
        "결재선 지정 완료",
        "결재요청 전송",
        "결재요청 확인대기",
        "처리 완료",
        "목록 유지",
        "결재요청 중지",
        "예외처리",
    }
)


def default_database_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    return base / "EAccAutomation" / "eacc.db"


class EAccApplication(tk.Tk):
    def __init__(self, repository: ImportRepository | None = None) -> None:
        super().__init__()
        self.title(f"{APP_TITLE} - 3단계 개발 테스트")
        self.geometry("1480x860")
        self.minsize(1100, 650)
        self.repository = repository or ImportRepository(default_database_path())
        self.browser_service = EAccountingBrowserService(
            default_database_path().parent / "downloads"
        )
        self.current_summary: ImportSummary | None = None
        self._rows_by_item: dict[str, ImportDisplayRow] = {}
        self._receipt_results: dict[str, ReceiptValidationResult] = {}
        self._account_results: dict[str, AccountValidationResult] = {}
        self._merchant_results: dict[str, MerchantLookupResult] = {}
        self._processing_statuses = self.repository.latest_processing_statuses(
            tuple(APPROVAL_RESULT_STATUSES)
        )
        self._current_target: UnsubmittedTransaction | None = None
        self._current_target_row: ImportDisplayRow | None = None
        self._pending_approval_transactions: dict[str, UnsubmittedTransaction] = {}
        self._excluded_transaction_ids: set[str] = set()
        self._bizno_client = BizNoLookupClient()
        self._pending_receipt_transaction: UnsubmittedTransaction | None = None
        self._pending_merchant_transaction: UnsubmittedTransaction | None = None
        self._automatic_target_validation = False
        self._batch_queue: list[UnsubmittedTransaction] = []
        self._batch_current: UnsubmittedTransaction | None = None
        self._batch_total = 0
        self._batch_completed = 0
        self._batch_login_credentials: LoginCredentials | None = None
        self._employee_names: tuple[str, ...] = ()
        self._employee_directory_error = ""
        try:
            self._employee_names = load_employee_names()
        except EmployeeDirectoryError as exc:
            # 특근자식비는 확인 가능한 직원 명단이 없으면 정상처리하지 않는다.
            self._employee_directory_error = str(exc)

        self._configure_styles()
        self._build_layout()
        # 비즈노 최초 연결은 느릴 수 있으므로 프로그램 화면을 막지 않고 미리 준비한다.
        self._bizno_client.prewarm()
        self._refresh_history()
        self._refresh_processing_history()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("맑은 고딕", 18, "bold"))
        style.configure("Subtitle.TLabel", font=("맑은 고딕", 10), foreground="#5B6470")
        style.configure("CardValue.TLabel", font=("맑은 고딕", 20, "bold"))
        style.configure("CardCaption.TLabel", font=("맑은 고딕", 9), foreground="#606A75")
        style.configure("Accent.TButton", font=("맑은 고딕", 10, "bold"), padding=(14, 8))
        style.configure("Treeview", rowheight=27, font=("맑은 고딕", 9))
        style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"))

    def _build_layout(self) -> None:
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True)

        sidebar = tk.Frame(root, bg="#263746", width=185)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(
            sidebar,
            text="E-Acc\n자동처리",
            bg="#263746",
            fg="white",
            font=("맑은 고딕", 17, "bold"),
            justify="left",
        ).pack(anchor="w", padx=22, pady=(28, 32))
        tk.Button(
            sidebar,
            text="미상신건수",
            bg="#3D79B8",
            fg="white",
            activebackground="#4B89C9",
            activeforeground="white",
            relief="flat",
            font=("맑은 고딕", 11, "bold"),
            anchor="w",
            padx=20,
            pady=12,
        ).pack(fill="x")
        tk.Button(
            sidebar,
            text="미처리건수  (추후 개발)",
            bg="#263746",
            fg="#93A3B3",
            disabledforeground="#93A3B3",
            relief="flat",
            state="disabled",
            font=("맑은 고딕", 10),
            anchor="w",
            padx=20,
            pady=12,
        ).pack(fill="x")
        tk.Label(
            sidebar,
            text="한 건씩 처리 · 1단계",
            bg="#263746",
            fg="#B9C5D0",
            font=("맑은 고딕", 9),
        ).pack(side="bottom", anchor="w", padx=20, pady=22)

        content = ttk.Frame(root, padding=(24, 18))
        content.pack(side="left", fill="both", expand=True)

        header = ttk.Frame(content)
        header.pack(fill="x")
        title_box = ttk.Frame(header)
        title_box.pack(side="left", fill="x", expand=True)
        ttk.Label(title_box, text="미상신내역 행 단위 검증", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            title_box,
            text="e-Acc에서 한 건씩 결재합니다. 프로그램은 선택 행의 검증·결과 이력·목록 재조회를 보조합니다.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        login_box = ttk.LabelFrame(title_box, text="i-NET 로그인 (현재 실행에서만 사용)", padding=(10, 6))
        login_box.pack(anchor="w", fill="x", pady=(12, 0), padx=(0, 12))
        self.login_id = tk.StringVar()
        self.login_password = tk.StringVar()
        ttk.Label(login_box, text="아이디").pack(side="left")
        self.login_id_entry = ttk.Entry(login_box, textvariable=self.login_id, width=17)
        self.login_id_entry.pack(side="left", padx=(6, 14))
        ttk.Label(login_box, text="비밀번호").pack(side="left")
        self.login_password_entry = ttk.Entry(
            login_box,
            textvariable=self.login_password,
            show="*",
            width=17,
        )
        self.login_password_entry.pack(side="left", padx=(6, 14))
        ttk.Label(
            login_box,
            text="입력값은 로그인 요청 직후 지워지며 저장하지 않습니다.",
            style="Subtitle.TLabel",
        ).pack(side="left")
        button_box = ttk.Frame(header)
        button_box.pack(side="right", padx=(16, 0))
        self.auto_import_button = ttk.Button(
            button_box,
            text="e-Acc 목록 불러오기",
            style="Accent.TButton",
            command=self._collect_from_eaccounting,
        )
        self.auto_import_button.pack(side="top", fill="x")
        self.manual_import_button = ttk.Button(
            button_box,
            text="다운로드 파일 직접 불러오기",
            command=self._choose_file,
        )
        self.manual_import_button.pack(side="top", fill="x", pady=(6, 0))
        self.current_target_button = ttk.Button(
            button_box,
            text="현재 첫 처리대상 읽기",
            command=self._read_current_first_target,
        )
        self.current_target_button.pack(side="top", fill="x", pady=(6, 0))

        self.current_target_message = tk.StringVar(value="현재 처리대상: 아직 읽지 않음")
        target_box = ttk.LabelFrame(content, text="현재 e-Acc 처리대상", padding=(12, 7))
        target_box.pack(fill="x", pady=(12, 0))
        ttk.Label(target_box, textvariable=self.current_target_message).pack(anchor="w")

        summary_frame = ttk.Frame(content)
        summary_frame.pack(fill="x", pady=(18, 12))
        self.summary_vars = {
            "원본 전체": tk.StringVar(value="0"),
            "신규": tk.StringVar(value="0"),
            "중복": tk.StringVar(value="0"),
            "오류": tk.StringVar(value="0"),
        }
        for index, (caption, variable) in enumerate(self.summary_vars.items()):
            card = ttk.LabelFrame(summary_frame, padding=(16, 9))
            card.grid(row=0, column=index, padx=(0, 10), sticky="ew")
            summary_frame.columnconfigure(index, weight=1)
            ttk.Label(card, textvariable=variable, style="CardValue.TLabel").pack(anchor="w")
            ttk.Label(card, text=caption, style="CardCaption.TLabel").pack(anchor="w")

        toolbar = ttk.Frame(content)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(toolbar, text="상태").pack(side="left")
        self.status_filter = tk.StringVar(value="전체")
        status_box = ttk.Combobox(
            toolbar,
            textvariable=self.status_filter,
            values=STATUS_FILTERS,
            state="readonly",
            width=9,
        )
        status_box.pack(side="left", padx=(6, 16))
        status_box.bind("<<ComboboxSelected>>", lambda _event: self._apply_filter())
        ttk.Label(toolbar, text="검색").pack(side="left")
        self.search_text = tk.StringVar()
        search_entry = ttk.Entry(toolbar, textvariable=self.search_text, width=30)
        search_entry.pack(side="left", padx=(6, 6))
        search_entry.bind("<Return>", lambda _event: self._apply_filter())
        ttk.Button(toolbar, text="적용", command=self._apply_filter).pack(side="left")
        ttk.Button(toolbar, text="초기화", command=self._reset_filter).pack(side="left", padx=(5, 0))
        self.receipt_button = ttk.Button(
            toolbar,
            text="선택 행 영수증 검증",
            command=self._download_selected_receipt,
        )
        self.receipt_button.pack(side="left", padx=(14, 0))
        self.batch_receipt_button = ttk.Button(
            toolbar,
            text="결재 후 목록 재조회",
            command=self._refresh_after_eacc_approval,
        )
        self.batch_receipt_button.pack(side="left", padx=(6, 0))
        self.approval_request_button = ttk.Button(
            toolbar,
            text="선택 행 결재요청",
            command=self._prepare_selected_approval_request,
        )
        self.approval_request_button.pack(side="left", padx=(6, 0))
        self.pg_lookup_button = ttk.Button(
            toolbar,
            text="선택 PG 상호 조회",
            command=self._lookup_selected_pg_merchant,
        )
        self.pg_lookup_button.pack(side="left", padx=(14, 0))
        self.actual_merchant_register_button = ttk.Button(
            toolbar,
            text="열린 실구매처 팝업 자동등록",
            command=self._register_selected_actual_merchant,
        )
        self.actual_merchant_register_button.pack(side="left", padx=(6, 0))
        self.file_label = ttk.Label(toolbar, text="불러온 파일 없음", style="Subtitle.TLabel")
        self.file_label.pack(side="right")

        self.progress = ttk.Progressbar(content, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 8))
        self.progress.pack_forget()

        self.notebook = ttk.Notebook(content)
        self.notebook.pack(fill="both", expand=True)
        transactions_tab = ttk.Frame(self.notebook)
        history_tab = ttk.Frame(self.notebook)
        processing_tab = ttk.Frame(self.notebook)
        self.notebook.add(transactions_tab, text="거래내역")
        self.notebook.add(history_tab, text="가져오기 이력")
        self.notebook.add(processing_tab, text="처리 결과 이력")

        self._build_transactions_table(transactions_tab)
        self._build_history_table(history_tab)
        self._build_processing_table(processing_tab)

        status_frame = ttk.Frame(content)
        status_frame.pack(fill="x", pady=(8, 0))
        self.status_message = tk.StringVar(value="미상신내역 파일을 선택해 주세요.")
        ttk.Label(status_frame, textvariable=self.status_message).pack(side="left")
        ttk.Label(
            status_frame,
            text=f"이력 DB: {self.repository.database_path}",
            style="Subtitle.TLabel",
        ).pack(side="right")

    def _build_transactions_table(self, parent: ttk.Frame) -> None:
        columns = (
            "행",
            "가져오기상태",
            "내부거래ID",
            "처리결과",
            "영수증판정",
            "영수증사유",
            "계정판정",
            "계정사유",
            "PG조회",
            "사업자번호",
            "비즈노상호",
            "실구매처등록",
            *DISPLAY_HEADERS,
            "오류사유",
        )
        self.transactions_tree = ttk.Treeview(parent, columns=columns, show="headings")
        vertical = ttk.Scrollbar(parent, orient="vertical", command=self.transactions_tree.yview)
        horizontal = ttk.Scrollbar(parent, orient="horizontal", command=self.transactions_tree.xview)
        self.transactions_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.transactions_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        widths = {
            "행": 55,
            "가져오기상태": 90,
            "내부거래ID": 175,
            "처리결과": 130,
            "영수증판정": 85,
            "영수증사유": 300,
            "계정판정": 85,
            "계정사유": 300,
            "PG조회": 100,
            "사업자번호": 110,
            "비즈노상호": 180,
            "실구매처등록": 120,
            "카드번호": 145,
            "카드소지자": 90,
            "승인번호": 95,
            "적요": 260,
            "거래처": 180,
            "사용금액": 100,
            "계정명": 170,
            "코스트센터": 150,
            "오류사유": 260,
        }
        for column in columns:
            self.transactions_tree.heading(column, text=column)
            anchor = "e" if column == "사용금액" else "center" if column in {"행", "가져오기상태"} else "w"
            self.transactions_tree.column(column, width=widths.get(column, 110), minwidth=55, anchor=anchor)
        self.transactions_tree.tag_configure("신규", background="#EFF8F1")
        self.transactions_tree.tag_configure("중복", background="#FFF7E8")
        self.transactions_tree.tag_configure("오류", background="#FDECEC")
        self.transactions_tree.tag_configure("현재", background="#E8F3FF")

    def _build_history_table(self, parent: ttk.Frame) -> None:
        columns = ("작업번호", "가져온시각", "파일명", "전체", "신규", "중복", "오류")
        self.history_tree = ttk.Treeview(parent, columns=columns, show="headings")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        widths = {"작업번호": 80, "가져온시각": 185, "파일명": 480, "전체": 80, "신규": 80, "중복": 80, "오류": 80}
        for column in columns:
            self.history_tree.heading(column, text=column)
            self.history_tree.column(column, width=widths[column], anchor="center" if column != "파일명" else "w")

    def _build_processing_table(self, parent: ttk.Frame) -> None:
        columns = ("시각", "상태", "승인번호", "증빙일자", "금액", "거래처", "사유", "내부거래ID")
        self.processing_tree = ttk.Treeview(parent, columns=columns, show="headings")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.processing_tree.yview)
        self.processing_tree.configure(yscrollcommand=scrollbar.set)
        self.processing_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        widths = {
            "시각": 175,
            "상태": 140,
            "승인번호": 100,
            "증빙일자": 105,
            "금액": 100,
            "거래처": 180,
            "사유": 440,
            "내부거래ID": 180,
        }
        for column in columns:
            self.processing_tree.heading(column, text=column)
            self.processing_tree.column(column, width=widths[column], anchor="center" if column not in {"거래처", "사유", "내부거래ID"} else "w")

    def _choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="미상신내역 파일 선택",
            filetypes=(("e-Accounting 미상신내역", "*.xls"), ("모든 파일", "*.*")),
        )
        if not path:
            return
        self._import_file(path)

    def _collect_from_eaccounting(self) -> None:
        credentials = self._consume_login_credentials()
        self._set_busy(True, "e-Accounting에서 미상신내역을 조회하고 다운로드하는 중입니다...")
        self.browser_service.collect_unsubmitted(
            credentials,
            on_success=lambda path: self.after(0, self._on_browser_download, path),
            on_error=lambda exc: self.after(0, self._on_browser_error, exc),
        )

    def _read_current_first_target(self) -> None:
        """Stage 1: get the first row currently visible in e-Acc, without approval."""
        self._set_busy(True, "e-Acc에서 코스트센터 검색 후 현재 첫 처리대상을 읽는 중입니다...")
        self.browser_service.read_current_first_target(
            self._consume_login_credentials(),
            on_success=lambda transaction: self.after(0, self._on_current_first_target, transaction),
            on_error=lambda exc: self.after(0, self._on_current_target_error, exc),
            excluded_transaction_ids=frozenset(self._excluded_transaction_ids),
        )

    def _on_current_first_target(self, transaction: UnsubmittedTransaction) -> None:
        self._set_busy(False)
        self._current_target = transaction
        # 한 건씩 처리 흐름에서는 이전에 내려받은 전체 Excel 목록이 현재 화면의
        # 첫 대상과 섞여 보이면 안 된다. 이 화면은 e-Acc에서 직접 읽은 현재 행만
        # 보여 주며, 다음 재검색에서 다시 현재 행을 갱신한다.
        self.current_summary = None
        self._receipt_results.clear()
        self._account_results.clear()
        self._merchant_results.clear()
        display_values = transaction.display_values()
        raw_values = tuple((*display_values[:19], "", *display_values[19:]))
        self._current_target_row = ImportDisplayRow(
            source_row_number=1,
            status="현재",
            transaction=transaction,
            raw_values=raw_values,
        )
        self._evaluate_account_rules(transaction, None)
        self.file_label.configure(text="e-Acc 현재 첫 행")
        self.summary_vars["원본 전체"].set("1")
        self.summary_vars["신규"].set("0")
        self.summary_vars["중복"].set("0")
        self.summary_vars["오류"].set("0")
        self.current_target_message.set(
            "현재 처리대상: "
            f"승인번호 {transaction.approval_number} / 증빙일자 {transaction.evidence_date} / "
            f"금액 {transaction.amount:,.0f} / 계정명 {transaction.account_name} / 업종 {transaction.business_type}"
        )
        self._record_processing_event(
            transaction,
            "현재 처리대상 읽음",
            "e-Acc 코스트센터 검색 결과의 현재 첫 행을 직접 읽었습니다.",
        )
        self._apply_filter()
        items = self.transactions_tree.get_children()
        if len(items) == 1:
            self.transactions_tree.selection_set(items[0])
            self.transactions_tree.focus(items[0])
        self.status_message.set("현재 첫 처리대상을 읽었습니다. 영수증 검증을 시작합니다...")
        self.after(50, self._validate_current_target_automatically, transaction)

    def _on_current_target_error(self, error: Exception) -> None:
        self._set_busy(False)
        if isinstance(error, NoEligibleTransactions):
            self._current_target = None
            self._current_target_row = None
            self.current_target_message.set("현재 처리대상: 없음 (이번 실행의 예외처리 행만 남음)")
            self.status_message.set("처리 가능 대상 없음: 예외처리한 행 외에는 남아 있지 않습니다.")
            return
        if isinstance(error, NoUnsubmittedTransactions):
            self._current_target = None
            self._current_target_row = None
            self.current_target_message.set("현재 처리대상: 없음 (미상신내역 검색 결과 0건)")
            self.status_message.set("처리 대상 없음: e-Acc 검색 결과가 0건입니다.")
            return
        self._on_browser_error(error)

    def _continue_after_exception(self, transaction: UnsubmittedTransaction) -> None:
        """Do not retry an exception row; re-read e-Acc for the next eligible row."""
        self._excluded_transaction_ids.add(transaction.transaction_id)
        self._set_busy(True, "예외 행을 제외하고 e-Acc의 다음 처리대상을 읽는 중입니다...")
        self.browser_service.read_current_first_target(
            None,
            on_success=lambda next_transaction: self.after(
                0, self._on_current_first_target, next_transaction
            ),
            on_error=lambda exc: self.after(0, self._on_current_target_error, exc),
            excluded_transaction_ids=frozenset(self._excluded_transaction_ids),
        )

    def _validate_current_target_automatically(
        self,
        transaction: UnsubmittedTransaction,
    ) -> None:
        """Stage 2: validate only the row that was just read from e-Acc."""
        if self._current_target is None or self._current_target.transaction_id != transaction.transaction_id:
            return
        account_result = self._account_results.get(transaction.transaction_id)
        if account_result is not None and account_result.status == "예외":
            self.status_message.set(f"계정별 예외처리: {account_result.reason_text}")
            self._continue_after_exception(transaction)
            return
        if transaction.evidence_status != "#":
            self._record_non_synchronized_evidence(transaction)
            self._record_processing_event(
                transaction,
                "처리 제외",
                "영수증 검증 대상이 아닌 증빙유무 상태입니다.",
            )
            self._continue_after_exception(transaction)
            return
        self._automatic_target_validation = True
        self._pending_receipt_transaction = transaction
        self._set_busy(True, "현재 첫 처리대상의 원본 영수증을 OCR 검증하는 중입니다...")
        self.browser_service.download_receipt(
            transaction,
            None,
            on_success=lambda result: self.after(0, self._on_receipt_download, result),
            on_error=lambda exc: self.after(0, self._on_receipt_error, exc),
        )

    def _consume_login_credentials(self) -> LoginCredentials | None:
        """Use typed credentials once, then remove them from the visible UI."""
        user_id = self.login_id.get().strip()
        password = self.login_password.get()
        self.login_id.set("")
        self.login_password.set("")
        if not user_id and not password:
            return None
        if not user_id or not password:
            # 현재 Edge 세션이 살아 있으면 로그인 정보가 없어도 기존 작업을 계속할 수
            # 있다. 세션이 없을 때는 브라우저 서비스가 한 번만 안내 메시지를 표시한다.
            return None
        return LoginCredentials(user_id=user_id, password=password)

    def _on_browser_download(self, path: Path) -> None:
        self._set_busy(False)
        self._import_file(str(path))

    def _on_browser_error(self, error: Exception) -> None:
        self._set_busy(False)
        if isinstance(error, NoUnsubmittedTransactions):
            self._show_empty_search_result()
            return
        if isinstance(error, BrowserLoginRequired):
            messagebox.showinfo("i-NET 로그인 필요", str(error), parent=self)
            self.status_message.set("i-NET 아이디와 비밀번호를 입력한 뒤 작업 버튼을 다시 눌러 주세요.")
            return
        if isinstance(error, BrowserAutomationError):
            messagebox.showerror("자동 불러오기 실패", str(error), parent=self)
        else:
            messagebox.showerror(
                "자동 불러오기 실패",
                f"e-Accounting 처리 중 예상하지 못한 오류가 발생했습니다.\n\n{error}",
                parent=self,
            )
        self.status_message.set(f"자동 불러오기 실패: {error}")

    def _show_empty_search_result(self) -> None:
        """Show a normal zero-row search result without treating it as an error."""
        self.current_summary = None
        self._receipt_results.clear()
        self._account_results.clear()
        self._merchant_results.clear()
        self._current_target = None
        self._current_target_row = None
        self._rows_by_item.clear()
        self.transactions_tree.delete(*self.transactions_tree.get_children())
        for variable in self.summary_vars.values():
            variable.set("0")
        self.file_label.configure(text="검색 결과 없음")
        self._reconcile_pending_approvals(())
        self.status_message.set("e-Accounting 미상신내역 검색 결과가 없습니다.")

    def _download_selected_receipt(self) -> None:
        selected = self.transactions_tree.selection()
        if len(selected) != 1:
            messagebox.showinfo(
                "영수증 행 선택",
                "거래내역에서 영수증을 확인할 행 하나를 선택해 주세요.",
                parent=self,
            )
            return
        row = self._rows_by_item.get(selected[0])
        if row is None or row.transaction is None:
            messagebox.showerror(
                "영수증 확인 불가",
                "오류 행은 영수증과 연결할 수 없습니다.",
                parent=self,
            )
            return
        if row.transaction.evidence_status != "#":
            self._record_non_synchronized_evidence(row.transaction)
            return
        self._pending_receipt_transaction = row.transaction
        self._set_busy(True, "선택 행의 원본 영수증을 찾고 다운로드하는 중입니다...")
        self.browser_service.download_receipt(
            row.transaction,
            self._consume_login_credentials(),
            on_success=lambda result: self.after(0, self._on_receipt_download, result),
            on_error=lambda exc: self.after(0, self._on_receipt_error, exc),
        )

    def _selected_transaction(self, action_name: str) -> UnsubmittedTransaction | None:
        selected = self.transactions_tree.selection()
        if len(selected) != 1:
            messagebox.showinfo(
                f"{action_name} 행 선택",
                "거래내역에서 한 행을 선택해 주세요.",
                parent=self,
            )
            return None
        row = self._rows_by_item.get(selected[0])
        if row is None or row.transaction is None:
            messagebox.showerror(
                f"{action_name} 불가",
                "오류 행은 처리할 수 없습니다.",
                parent=self,
            )
            return None
        return row.transaction

    def _record_processing_event(
        self,
        transaction: UnsubmittedTransaction,
        status: str,
        reason: str = "",
    ) -> ProcessingEvent:
        event = self.repository.record_processing_event(transaction, status, reason)
        if status in APPROVAL_RESULT_STATUSES:
            self._processing_statuses[transaction.transaction_id] = status
            # 사용자가 보는 거래내역의 처리결과도 즉시 같은 값으로 갱신한다.
            self._apply_filter()
        self._refresh_processing_history()
        return event

    def _refresh_after_eacc_approval(self) -> None:
        """Re-read e-Acc only after the user has requested approval for one row."""
        transaction = self._selected_transaction("결재 후 목록 재조회")
        if transaction is None:
            return
        self._pending_approval_transactions[transaction.transaction_id] = transaction
        self._record_processing_event(
            transaction,
            "결재요청 확인대기",
            "e-Acc에서 이 행의 결재 요청을 실행한 뒤 최신 목록으로 확인을 시작했습니다.",
        )
        self._set_busy(True, "e-Acc의 최신 목록을 다시 조회하여 결재 요청 결과를 확인하는 중입니다...")
        self.browser_service.collect_unsubmitted(
            self._consume_login_credentials(),
            on_success=lambda path: self.after(0, self._on_browser_download, path),
            on_error=lambda exc: self.after(0, self._on_browser_error, exc),
        )

    def _approval_request_block_reason(
        self,
        transaction: UnsubmittedTransaction,
    ) -> str | None:
        """Return the reason an unverified row must not reach e-Acc approval."""
        account_result = self._account_results.get(transaction.transaction_id)
        if account_result is None:
            receipt = self._receipt_results.get(transaction.transaction_id)
            account_result = self._evaluate_account_rules(
                transaction,
                receipt.ocr_text if receipt is not None else None,
            )
        if account_result.status == "예외":
            return account_result.reason_text
        if account_result.status == "검증대기":
            return "계정별 영수증 검증 결과가 없습니다. 먼저 영수증 검증을 완료해 주세요."
        if transaction.evidence_status != "#":
            return (
                f"증빙유무가 '{transaction.evidence_status}'입니다. "
                "영수증 동기화 완료(#) 상태가 아니므로 결재요청할 수 없습니다."
            )
        validation = self._receipt_results.get(transaction.transaction_id)
        if validation is None:
            return "현재 행의 영수증 OCR 검증 결과가 없습니다. 먼저 영수증 검증을 완료해 주세요."
        if validation.status != "정상":
            return "영수증 OCR 판정이 정상 상태가 아니므로 결재요청할 수 없습니다."
        if is_pg_business_type(transaction.business_type) and not transaction.actual_merchant_name.strip():
            merchant = self._merchant_results.get(transaction.transaction_id)
            if merchant is None or merchant.registration_status != "등록 완료":
                return (
                    "PG일반 행의 실구매처 등록이 완료되지 않았습니다. "
                    "사업자번호·비즈노 상호조회와 실구매처 등록을 먼저 완료해 주세요."
                )
        return None

    def _evaluate_account_rules(
        self,
        transaction: UnsubmittedTransaction,
        receipt_text: str | None,
    ) -> AccountValidationResult:
        """Apply account rules and persist a blocking outcome exactly once."""
        employee_names = self._employee_names
        result = validate_account_rules(transaction, employee_names, receipt_text)
        # No usable employee master must never make an overtime meal pass by
        # accident.  The standard insufficient-headcount reason is retained.
        if not employee_names and transaction.account_name.strip() == "특근자식비":
            result = AccountValidationResult(
                transaction_id=transaction.transaction_id,
                status="예외",
                reasons=(
                    "특근자식비 사용인원 불충족"
                    f" (직원 명단 확인 불가: {self._employee_directory_error or '명단 없음'})",
                ),
            )
        self._account_results[transaction.transaction_id] = result
        if (
            result.status == "예외"
            and self._processing_statuses.get(transaction.transaction_id) != "예외처리"
        ):
            self._record_processing_event(transaction, "예외처리", result.reason_text)
        return result

    def _prepare_selected_approval_request(self) -> None:
        """Open the one-row approval line; the actual request remains user-gated."""
        transaction = self._selected_transaction("결재요청")
        if transaction is None:
            return
        blocked_reason = self._approval_request_block_reason(transaction)
        if blocked_reason is not None:
            messagebox.showwarning("결재요청 불가", blocked_reason, parent=self)
            self.status_message.set(f"결재요청 차단: {blocked_reason}")
            return
        self._set_busy(True, "검증 정상인 현재 한 건을 선택하고 결재선 지정 창을 여는 중입니다...")
        self.browser_service.open_approval_line(
            transaction,
            self._consume_login_credentials(),
            on_success=lambda message: self.after(
                0, self._on_approval_line_opened, transaction, message
            ),
            on_error=lambda exc: self.after(
                0, self._on_approval_request_error, transaction, exc
            ),
        )

    def _on_approval_line_opened(
        self,
        transaction: UnsubmittedTransaction,
        message: str,
    ) -> None:
        self._record_processing_event(transaction, "결재선 지정 완료", message)
        self._set_busy(True, "e-Acc 결재선 지정 후 결재요청을 전송하는 중입니다...")
        self.browser_service.submit_prepared_approval(
            transaction,
            on_success=lambda result: self.after(
                0, self._on_approval_request_submitted, transaction, result
            ),
            on_error=lambda exc: self.after(
                0, self._on_approval_request_error, transaction, exc
            ),
        )

    def _on_approval_request_cancelled(
        self,
        transaction: UnsubmittedTransaction,
        result: str,
    ) -> None:
        self._set_busy(False)
        self._record_processing_event(transaction, "결재요청 중지", result)
        self.status_message.set(result)

    def _on_approval_request_submitted(
        self,
        transaction: UnsubmittedTransaction,
        result: str,
    ) -> None:
        self._set_busy(False)
        self._pending_approval_transactions[transaction.transaction_id] = transaction
        self._record_processing_event(transaction, "결재요청 전송", result)
        self.status_message.set(
            "결재요청을 전송했습니다. '결재 후 목록 재조회' 버튼으로 처리 결과를 확인해 주세요."
        )

    def _on_approval_request_error(
        self,
        transaction: UnsubmittedTransaction,
        error: Exception,
    ) -> None:
        self._set_busy(False)
        if isinstance(error, ApprovalInProgressError):
            reason = "다른 사용자가 처리중"
            self._record_processing_event(transaction, "예외처리", reason)
            self.status_message.set(
                "예외처리: 다른 사용자가 이미 처리 중인 행입니다. "
                "결재선 지정 창과 선택 상태를 정리했습니다."
            )
            self._continue_after_exception(transaction)
            return
        self._on_browser_error(error)

    def _lookup_selected_pg_merchant(self) -> None:
        transaction = self._selected_transaction("PG 상호 조회")
        if transaction is None:
            return
        if not is_pg_business_type(transaction.business_type):
            messagebox.showinfo(
                "PG 조회 대상 아님",
                f"업종 '{transaction.business_type}'은(는) PG일반 조회 대상이 아닙니다.",
                parent=self,
            )
            return
        if transaction.evidence_status != "#":
            status, reason = self._non_synchronized_evidence_status(transaction)
            self._merchant_results[transaction.transaction_id] = MerchantLookupResult(
                transaction_id=transaction.transaction_id,
                status=status,
                reason=reason,
            )
            self._apply_filter()
            return
        self._pending_merchant_transaction = transaction
        self._set_busy(True, "PG 영수증에서 사업자번호를 확인하는 중입니다...")
        self.browser_service.download_receipt(
            transaction,
            self._consume_login_credentials(),
            on_success=lambda result: self.after(0, self._on_pg_receipt_download, result),
            on_error=lambda exc: self.after(0, self._on_pg_merchant_error, exc),
        )

    def _on_pg_receipt_download(self, result: ReceiptImageResult) -> None:
        transaction = self._pending_merchant_transaction
        if transaction is None or transaction.transaction_id != result.transaction_id:
            self._on_pg_merchant_error(RuntimeError("PG 조회 거래 연결을 찾지 못했습니다."))
            return
        self.status_message.set("PG 영수증의 사업자번호를 OCR로 확인하는 중입니다...")

        def recognize() -> None:
            try:
                validation = validate_receipt_images(transaction, result.image_paths)
                extraction = extract_business_numbers(validation.ocr_text)
                if len(extraction.candidates) != 1:
                    raise MerchantLookupError(extraction.reason)
                business_number = extraction.candidates[0]
            except Exception as exc:
                self.after(0, self._on_pg_merchant_error, exc)
                return
            self.after(0, self._start_moneypin_lookup, transaction, validation, business_number)

        threading.Thread(target=recognize, name="pg-business-number-ocr", daemon=True).start()

    def _start_moneypin_lookup(
        self,
        transaction: UnsubmittedTransaction,
        validation: ReceiptValidationResult,
        business_number: str,
    ) -> None:
        self._receipt_results[transaction.transaction_id] = validation
        self._evaluate_account_rules(transaction, validation.ocr_text)
        self._apply_filter()
        self.status_message.set("비즈노에서 사업자번호로 상호를 조회하는 중입니다...")

        def lookup() -> None:
            try:
                merchant_name = self._bizno_client.lookup(business_number)
            except Exception as exc:
                self.after(0, self._on_pg_merchant_error, exc)
            else:
                self.after(
                    0,
                    self._on_pg_merchant_lookup,
                    transaction,
                    business_number,
                    merchant_name,
                )

        threading.Thread(target=lookup, name="moneypin-merchant-lookup", daemon=True).start()

    def _on_pg_merchant_lookup(
        self,
        transaction: UnsubmittedTransaction,
        business_number: str,
        merchant_name: str,
    ) -> None:
        self._set_busy(False)
        self._pending_merchant_transaction = None
        self._merchant_results[transaction.transaction_id] = MerchantLookupResult(
            transaction_id=transaction.transaction_id,
            status="조회완료",
            business_number=business_number,
            merchant_name=merchant_name,
        )
        self._record_processing_event(
            transaction,
            "PG 상호조회 완료",
            f"사업자번호 {business_number} / 비즈노 상호 {merchant_name}",
        )
        self._apply_filter()
        self.status_message.set(
            f"PG 상호 조회 완료: 사업자번호 {business_number} / 비즈노 상호 {merchant_name}"
        )

    def _on_pg_merchant_error(self, error: Exception) -> None:
        self._set_busy(False)
        transaction = self._pending_merchant_transaction
        self._pending_merchant_transaction = None
        if transaction is not None:
            reason = str(error)
            self._merchant_results[transaction.transaction_id] = MerchantLookupResult(
                transaction_id=transaction.transaction_id,
                status="확인 필요",
                reason=reason,
            )
            self._record_processing_event(transaction, "PG 상호조회 확인 필요", reason)
            self._apply_filter()
        self.status_message.set(f"PG 상호 조회 확인 필요: {error}")

    def _register_selected_actual_merchant(self) -> None:
        transaction = self._selected_transaction("실구매처 등록")
        if transaction is None:
            return
        result = self._merchant_results.get(transaction.transaction_id)
        if result is None or result.status != "조회완료":
            messagebox.showinfo(
                "상호 조회 필요",
                "먼저 해당 PG일반 행의 '선택 PG 상호 조회'를 완료해 주세요.",
                parent=self,
            )
            return
        if transaction.actual_merchant_name.strip():
            messagebox.showinfo(
                "등록 대상 아님",
                f"이미 실구매처명 '{transaction.actual_merchant_name}'이 등록된 행입니다.",
                parent=self,
            )
            return
        confirm = messagebox.askyesno(
            "실구매처 등록 확인",
            "아래 값을 e-Accounting에 등록합니다.\n\n"
            f"사업자번호: {result.business_number}\n"
            f"사업자명: {result.merchant_name}\n\n"
            "등록할까요?",
            parent=self,
        )
        if not confirm:
            return
        self._set_busy(True, "e-Accounting 실구매처 등록 팝업에 조회 결과를 등록하는 중입니다...")
        self.browser_service.register_actual_merchant(
            transaction,
            result.business_number,
            result.merchant_name,
            self._consume_login_credentials(),
            on_success=lambda message: self.after(0, self._on_actual_merchant_registered, transaction, message),
            on_error=lambda exc: self.after(0, self._on_actual_merchant_registration_error, transaction, exc),
        )

    def _on_actual_merchant_registered(
        self,
        transaction: UnsubmittedTransaction,
        message: str,
    ) -> None:
        self._set_busy(False)
        result = self._merchant_results.get(transaction.transaction_id)
        if result is not None:
            self._merchant_results[transaction.transaction_id] = replace(
                result,
                registration_status="등록 완료",
                registration_reason=message,
            )
        self._record_processing_event(transaction, "실구매처 등록 완료", message)
        self._apply_filter()
        self.status_message.set(f"실구매처 등록 완료: {message}")

    def _on_actual_merchant_registration_error(
        self,
        transaction: UnsubmittedTransaction,
        error: Exception,
    ) -> None:
        self._set_busy(False)
        result = self._merchant_results.get(transaction.transaction_id)
        if result is not None:
            self._merchant_results[transaction.transaction_id] = replace(
                result,
                registration_status="등록 확인 필요",
                registration_reason=str(error),
            )
        self._record_processing_event(transaction, "실구매처 등록 확인 필요", str(error))
        self._apply_filter()
        title = "실구매처 등록 실패" if isinstance(error, ActualMerchantRegistrationError) else "실구매처 등록 오류"
        messagebox.showerror(title, str(error), parent=self)
        self.status_message.set(f"실구매처 등록 확인 필요: {error}")

    def _on_receipt_download(self, result: ReceiptImageResult) -> None:
        transaction = next(
            (
                row.transaction
                for row in self._rows_by_item.values()
                if row.transaction is not None
                and row.transaction.transaction_id == result.transaction_id
            ),
            None,
        )
        if transaction is None:
            self._on_receipt_error(RuntimeError("선택 거래 정보를 다시 찾지 못했습니다."))
            return
        self.status_message.set("원본 영수증의 승인번호·증빙일자·사용금액을 OCR로 확인하는 중입니다...")

        def validate() -> None:
            try:
                validation = validate_receipt_images(transaction, result.image_paths)
            except Exception as exc:
                self.after(0, self._on_receipt_error, exc)
            else:
                self.after(0, self._on_receipt_validation, result, validation)

        threading.Thread(target=validate, name="receipt-ocr-worker", daemon=True).start()

    def _on_receipt_validation(
        self,
        images: ReceiptImageResult,
        validation: ReceiptValidationResult,
    ) -> None:
        automatic = self._automatic_target_validation
        self._automatic_target_validation = False
        self._set_busy(False)
        self._receipt_results[validation.transaction_id] = validation
        transaction = next(
            (
                row.transaction
                for row in self._rows_by_item.values()
                if row.transaction is not None and row.transaction.transaction_id == validation.transaction_id
            ),
            None,
        )
        if transaction is not None:
            self._evaluate_account_rules(transaction, validation.ocr_text)
            details = " / ".join(
                f"{check.field_name} {'일치' if check.is_match else '불일치'}"
                for check in validation.checks
            )
            self._record_processing_event(transaction, f"영수증 {validation.status}", details)
        self._pending_receipt_transaction = None
        self._apply_filter()
        details = " / ".join(
            f"{check.field_name} {'일치' if check.is_match else '불일치'}"
            for check in validation.checks
        )
        direction = f"회전 {validation.rotation_degrees}°"
        if validation.orientation_ambiguous:
            direction = validation.orientation_reason
        self.status_message.set(f"영수증 판정 {validation.status}: {details} / {direction}")
        if automatic:
            account_result = (
                self._account_results.get(transaction.transaction_id)
                if transaction is not None
                else None
            )
            if account_result is not None and account_result.status == "예외":
                self.status_message.set(f"계정별 예외처리: {account_result.reason_text}")
            elif validation.status == "정상" and transaction is not None and is_pg_business_type(transaction.business_type):
                self._start_pg_lookup_from_validation(transaction, validation)
            elif validation.status != "정상" and transaction is not None:
                self._record_processing_event(transaction, "처리 제외", details or validation.status)
                self._continue_after_exception(transaction)
            return
        self._show_receipt_viewer(images, validation)

    def _on_receipt_error(self, error: Exception) -> None:
        automatic = self._automatic_target_validation
        self._automatic_target_validation = False
        self._set_busy(False)
        transaction = self._pending_receipt_transaction
        self._pending_receipt_transaction = None
        if isinstance(error, ReceiptNotAvailable):
            title = "영수증 없음"
            # 영수증이 없는 사실은 안내창으로 끝내지 않고, 해당 행의 영수증판정/사유에도
            # 남긴다. #인데 이미지가 없으면 동기화 상태가 변할 수 있으므로,
            # 실제 영수증 미등록으로 단정하지 않고 재조회 대상으로 분류한다.
            if transaction is not None:
                if transaction.evidence_status == "#":
                    status = "재조회 필요"
                    reason = (
                        "증빙유무 #이나 영수증 이미지가 확인되지 않습니다. "
                        f"미상신내역을 다시 불러와 동기화 상태를 확인해 주세요. ({error})"
                    )
                else:
                    status, reason = self._non_synchronized_evidence_status(transaction)
                self._receipt_results[transaction.transaction_id] = ReceiptValidationResult(
                    transaction_id=transaction.transaction_id,
                    status=status,
                    checks=(),
                    ocr_text="",
                    reason=reason,
                )
                self._record_processing_event(transaction, f"영수증 {status}", reason)
                if automatic:
                    self._record_processing_event(transaction, "처리 제외", reason)
                self._apply_filter()
            if automatic:
                self.status_message.set(f"{title}: {error}")
                if transaction is not None:
                    self._continue_after_exception(transaction)
                return
            messagebox.showinfo(title, str(error), parent=self)
            self.status_message.set(f"{title}: {error}")
            return

        title = "영수증 가져오기 실패"
        if transaction is not None:
            self._record_processing_event(transaction, "영수증 판독불가", str(error))
            if automatic:
                self._record_processing_event(transaction, "처리 제외", str(error))
        if not automatic:
            messagebox.showerror(title, str(error), parent=self)
        elif transaction is not None:
            self._continue_after_exception(transaction)
        self.status_message.set(f"{title}: {error}")

    @staticmethod
    def _non_synchronized_evidence_status(
        transaction: UnsubmittedTransaction,
    ) -> tuple[str, str]:
        if transaction.evidence_status == "Y":
            return (
                "동기화 대기",
                "사용자가 증빙을 등록했으며 e-Accounting 동기화를 기다리는 상태입니다.",
            )
        if transaction.evidence_status == "N":
            return (
                "영수증 미등록",
                "사용자가 증빙을 등록하지 않은 상태입니다. 영수증 등록 또는 영수증 불필요 처리 대상입니다.",
            )
        return ("검증 보류", f"증빙유무 {transaction.evidence_status}: 검증 대상 상태가 아닙니다.")

    def _record_non_synchronized_evidence(
        self,
        transaction: UnsubmittedTransaction,
    ) -> None:
        status, reason = self._non_synchronized_evidence_status(transaction)
        self._receipt_results[transaction.transaction_id] = ReceiptValidationResult(
            transaction_id=transaction.transaction_id,
            status=status,
            checks=(),
            ocr_text="",
            reason=reason,
        )
        self._record_processing_event(transaction, f"영수증 {status}", reason)
        self._apply_filter()
        self.status_message.set(f"영수증 판정 {status}: {reason}")

    def _start_pg_lookup_from_validation(
        self,
        transaction: UnsubmittedTransaction,
        validation: ReceiptValidationResult,
    ) -> None:
        """Reuse the validated receipt OCR text for the PG business-number lookup."""
        self._pending_merchant_transaction = transaction
        try:
            extraction = extract_business_numbers(validation.ocr_text)
            if len(extraction.candidates) != 1:
                raise MerchantLookupError(extraction.reason)
            business_number = extraction.candidates[0]
        except Exception as exc:
            self._on_pg_merchant_error(exc)
            return
        self._set_busy(True, "PG 영수증의 사업자번호를 비즈노에서 조회하는 중입니다...")
        self._start_moneypin_lookup(transaction, validation, business_number)

    def _validate_all_synchronized_receipts(self) -> None:
        if self.current_summary is None:
            messagebox.showinfo(
                "미상신내역 불러오기 필요",
                "먼저 최신 미상신내역을 불러온 후 # 영수증 전체 검증을 실행해 주세요.",
                parent=self,
            )
            return
        transactions = [
            row.transaction
            for row in self.current_summary.rows
            if row.transaction is not None and row.transaction.evidence_status == "#"
        ]
        if not transactions:
            messagebox.showinfo("검증 대상 없음", "현재 조회 결과에 # 상태의 영수증이 없습니다.", parent=self)
            return
        self._batch_queue = transactions
        self._batch_total = len(transactions)
        self._batch_completed = 0
        self._batch_current = None
        self._batch_login_credentials = self._consume_login_credentials()
        self._set_busy(True, f"# 영수증 전체 OCR 검증 준비: {self._batch_total:,}건")
        self._process_next_batch_receipt()

    def _process_next_batch_receipt(self) -> None:
        if not self._batch_queue:
            self._finish_batch_receipt_validation()
            return
        transaction = self._batch_queue.pop(0)
        self._batch_current = transaction
        self.status_message.set(
            f"# 영수증 OCR 검증 중: {self._batch_completed + 1:,}/{self._batch_total:,}건 "
            f"(승인번호 {transaction.approval_number})"
        )
        credentials = self._batch_login_credentials
        self._batch_login_credentials = None
        self.browser_service.download_receipt(
            transaction,
            credentials,
            on_success=lambda result: self.after(0, self._on_batch_receipt_download, result),
            on_error=lambda exc: self.after(0, self._on_batch_receipt_error, exc),
        )

    def _on_batch_receipt_download(self, result: ReceiptImageResult) -> None:
        transaction = self._batch_current
        if transaction is None or transaction.transaction_id != result.transaction_id:
            self._on_batch_receipt_error(RuntimeError("일괄 검증 거래 연결을 찾지 못했습니다."))
            return

        def validate() -> None:
            try:
                validation = validate_receipt_images(transaction, result.image_paths)
            except Exception as exc:
                self.after(0, self._on_batch_receipt_error, exc)
            else:
                self.after(0, self._on_batch_receipt_validation, validation)

        threading.Thread(target=validate, name="receipt-ocr-batch-worker", daemon=True).start()

    def _on_batch_receipt_validation(self, validation: ReceiptValidationResult) -> None:
        self._receipt_results[validation.transaction_id] = validation
        transaction = self._batch_current
        if transaction is not None and transaction.transaction_id == validation.transaction_id:
            self._evaluate_account_rules(transaction, validation.ocr_text)
        self._finish_one_batch_receipt()

    def _on_batch_receipt_error(self, error: Exception) -> None:
        transaction = self._batch_current
        if transaction is not None:
            if isinstance(error, ReceiptNotAvailable):
                status = "재조회 필요"
                reason = (
                    "증빙유무 #이나 영수증 이미지가 확인되지 않습니다. "
                    f"미상신내역을 다시 불러와 동기화 상태를 확인해 주세요. ({error})"
                )
            else:
                status = "판독불가"
                reason = f"영수증 다운로드 또는 OCR 처리 실패: {error}"
            self._receipt_results[transaction.transaction_id] = ReceiptValidationResult(
                transaction_id=transaction.transaction_id,
                status=status,
                checks=(),
                ocr_text="",
                reason=reason,
            )
        self._finish_one_batch_receipt()

    def _finish_one_batch_receipt(self) -> None:
        self._batch_completed += 1
        self._batch_current = None
        self._apply_filter()
        self.after(50, self._process_next_batch_receipt)

    def _finish_batch_receipt_validation(self) -> None:
        result_counts: dict[str, int] = {}
        for result in self._receipt_results.values():
            result_counts[result.status] = result_counts.get(result.status, 0) + 1
        summary = ", ".join(
            f"{status} {count:,}건" for status, count in sorted(result_counts.items())
        )
        self._batch_current = None
        self._set_busy(False)
        self._apply_filter()
        self.status_message.set(f"# 영수증 전체 OCR 검증 완료: {self._batch_completed:,}건 / {summary}")
        messagebox.showinfo(
            "# 영수증 전체 OCR 검증 완료",
            f"검증 {self._batch_completed:,}건이 완료되었습니다.\n\n{summary}",
            parent=self,
        )

    def _show_receipt_viewer(
        self,
        result: ReceiptImageResult,
        validation: ReceiptValidationResult,
    ) -> None:
        viewer = tk.Toplevel(self)
        viewer.title(f"영수증 {validation.status} - {result.transaction_id}")
        viewer.geometry("960x840")
        viewer.minsize(600, 500)
        ttk.Label(
            viewer,
            text=f"원본 {len(result.image_paths)}장 · DocIRN {', '.join(result.doc_irns)} · CorpNo {result.corp_no}",
            style="Subtitle.TLabel",
        ).pack(fill="x", padx=12, pady=(10, 6))
        result_box = ttk.LabelFrame(viewer, text=f"OCR 판정: {validation.status}", padding=(10, 6))
        result_box.pack(fill="x", padx=12, pady=(0, 8))
        ttk.Label(
            result_box,
            text=(
                validation.orientation_reason
                if validation.orientation_ambiguous
                else f"✓ 이미지 방향: OCR 적용 회전각도 {validation.rotation_degrees}°"
            ),
        ).pack(anchor="w")
        for check in validation.checks:
            ttk.Label(
                result_box,
                text=(
                    f"{'✓' if check.is_match else '✕'} {check.field_name}: "
                    f"화면 {check.expected_value} / OCR {check.detected_value or '찾지 못함'}"
                    + ("" if check.is_match else f" — {check.reason}")
                ),
            ).pack(anchor="w")
        notebook = ttk.Notebook(viewer)
        notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        viewer._receipt_photos = []  # type: ignore[attr-defined]
        for index, path in enumerate(result.image_paths, start=1):
            tab = ttk.Frame(notebook)
            notebook.add(tab, text=f"영수증 {index}")
            image = Image.open(path)
            if validation.rotation_degrees:
                image = image.rotate(validation.rotation_degrees, expand=True)
            image.thumbnail((840, 650), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            viewer._receipt_photos.append(photo)  # type: ignore[attr-defined]
            ttk.Label(tab, image=photo, anchor="center").pack(fill="both", expand=True)
            ttk.Label(tab, text=str(path), style="Subtitle.TLabel").pack(fill="x", padx=8, pady=6)
        ocr_tab = ttk.Frame(notebook)
        notebook.add(ocr_tab, text="OCR 원문")
        ocr_text = tk.Text(ocr_tab, wrap="word", font=("맑은 고딕", 10))
        ocr_text.pack(fill="both", expand=True, padx=8, pady=8)
        ocr_text.insert("1.0", validation.ocr_text)
        ocr_text.configure(state="disabled")

    def _set_busy(self, busy: bool, message: str = "") -> None:
        state = "disabled" if busy else "normal"
        self.auto_import_button.configure(state=state)
        self.manual_import_button.configure(state=state)
        self.receipt_button.configure(state=state)
        self.batch_receipt_button.configure(state=state)
        self.approval_request_button.configure(state=state)
        self.pg_lookup_button.configure(state=state)
        self.actual_merchant_register_button.configure(state=state)
        self.current_target_button.configure(state=state)
        self.login_id_entry.configure(state=state)
        self.login_password_entry.configure(state=state)
        if busy:
            self.progress.pack(fill="x", pady=(0, 8), before=self.notebook)
            self.progress.start(10)
            if message:
                self.status_message.set(message)
        else:
            self.progress.stop()
            self.progress.pack_forget()

    def _import_file(self, path: str) -> None:
        self._set_busy(True, "파일을 읽고 중복 여부를 확인하는 중입니다...")
        self.update_idletasks()
        try:
            summary = self.repository.import_file(path)
        except (OSError, WorkbookFormatError) as exc:
            messagebox.showerror("불러오기 실패", str(exc), parent=self)
            self.status_message.set(f"불러오기 실패: {exc}")
            return
        except Exception as exc:  # UI 경계에서 예상하지 못한 오류를 행 작업과 분리한다.
            messagebox.showerror("예상하지 못한 오류", f"파일 처리 중 오류가 발생했습니다.\n\n{exc}", parent=self)
            self.status_message.set("예상하지 못한 오류로 불러오기에 실패했습니다.")
            return
        finally:
            self._set_busy(False)

        self.current_summary = summary
        # 증빙유무는 동기화 진행에 따라 Y/#/N으로 바뀔 수 있으므로, 최신 조회 결과가
        # 들어오면 이전 OCR·재조회 결과를 보존하지 않는다.
        self._receipt_results.clear()
        self._account_results.clear()
        self._merchant_results.clear()
        for row in summary.rows:
            if row.transaction is not None:
                self._evaluate_account_rules(row.transaction, None)
        self.file_label.configure(text=Path(summary.source_file).name)
        self.summary_vars["원본 전체"].set(f"{summary.total_count:,}")
        self.summary_vars["신규"].set(f"{summary.new_count:,}")
        self.summary_vars["중복"].set(f"{summary.duplicate_count:,}")
        self.summary_vars["오류"].set(f"{summary.error_count:,}")
        self._apply_filter()
        self._refresh_history()
        self._reconcile_pending_approvals(summary.rows)
        self.status_message.set(
            f"작업 #{summary.job_id} 완료: 전체 {summary.total_count:,}건 / "
            f"신규 {summary.new_count:,}건 / 중복 {summary.duplicate_count:,}건 / 오류 {summary.error_count:,}건"
        )

    def _apply_filter(self) -> None:
        self.transactions_tree.delete(*self.transactions_tree.get_children())
        self._rows_by_item.clear()
        if self.current_summary is None and self._current_target_row is None:
            return
        rows = (
            self.current_summary.rows
            if self.current_summary is not None
            else (self._current_target_row,)
        )
        selected_status = self.status_filter.get()
        keyword = self.search_text.get().strip().casefold()
        visible_count = 0
        for row in rows:
            if row is None:
                continue
            if selected_status != "전체" and row.status != selected_status:
                continue
            searchable = " ".join(row.raw_values).casefold()
            if keyword and keyword not in searchable and keyword not in row.error_message.casefold():
                continue
            values = self._tree_values(row)
            item_id = self.transactions_tree.insert("", "end", values=values, tags=(row.status,))
            self._rows_by_item[item_id] = row
            visible_count += 1
        self.status_message.set(f"현재 {visible_count:,}건 표시 중")

    def _tree_values(self, row: ImportDisplayRow) -> tuple[str, ...]:
        if row.transaction is not None:
            receipt = self._receipt_results.get(row.transaction.transaction_id)
            if receipt is None and row.transaction.evidence_status != "#":
                evidence_status, evidence_reason = self._non_synchronized_evidence_status(
                    row.transaction
                )
                receipt = ReceiptValidationResult(
                    transaction_id=row.transaction.transaction_id,
                    status=evidence_status,
                    checks=(),
                    ocr_text="",
                    reason=evidence_reason,
                )
            receipt_reason = ""
            if receipt is not None:
                receipt_reason = " / ".join(
                    check.reason for check in receipt.checks if not check.is_match
                )
                if receipt.reason:
                    receipt_reason = " / ".join(filter(None, (receipt_reason, receipt.reason)))
                if receipt.orientation_reason:
                    receipt_reason = " / ".join(filter(None, (receipt_reason, receipt.orientation_reason)))
            merchant = self._merchant_results.get(row.transaction.transaction_id)
            account = self._account_results.get(row.transaction.transaction_id)
            processing_result = self._processing_statuses.get(
                row.transaction.transaction_id, "미처리"
            )
            if account is not None and account.status == "예외":
                processing_result = "예외처리"
            return (
                row.source_row_number,
                row.status,
                row.transaction.transaction_id,
                processing_result,
                receipt.status if receipt is not None else "미검증",
                receipt_reason,
                account.status if account is not None else "미검증",
                account.reason_text if account is not None else "",
                merchant.status if merchant is not None else "대상 아님" if not is_pg_business_type(row.transaction.business_type) else "미조회",
                merchant.business_number if merchant is not None else "",
                merchant.merchant_name if merchant is not None else "",
                merchant.registration_status if merchant is not None else "",
                *row.transaction.display_values(),
                row.error_message,
            )
        raw = list(row.raw_values)
        if len(raw) < 21:
            raw.extend([""] * (21 - len(raw)))
        business_values = [value for index, value in enumerate(raw[:21]) if index != 19]
        return (
            row.source_row_number,
            row.status,
            "",
            "",
            "검증불가",
            row.error_message,
            "",
            "",
            "",
            "",
            "",
            "",
            *business_values,
            row.error_message,
        )

    def _reset_filter(self) -> None:
        self.status_filter.set("전체")
        self.search_text.set("")
        self._apply_filter()

    def _refresh_history(self) -> None:
        self.history_tree.delete(*self.history_tree.get_children())
        for item in self.repository.recent_history():
            self.history_tree.insert(
                "",
                "end",
                values=(
                    item.job_id,
                    item.imported_at.replace("T", " "),
                    Path(item.source_file).name,
                    item.total_count,
                    item.new_count,
                    item.duplicate_count,
                    item.error_count,
                ),
            )

    def _refresh_processing_history(self) -> None:
        self.processing_tree.delete(*self.processing_tree.get_children())
        for event in self.repository.recent_processing_events():
            self.processing_tree.insert(
                "",
                "end",
                values=(
                    event.event_at.replace("T", " "),
                    event.status,
                    event.approval_number,
                    event.evidence_date,
                    event.amount,
                    event.merchant,
                    event.reason,
                    event.transaction_id,
                ),
            )

    def _reconcile_pending_approvals(
        self,
        rows: tuple[ImportDisplayRow, ...],
    ) -> None:
        """Record the result of one-row e-Acc approval after a fresh list read."""
        if not self._pending_approval_transactions:
            return
        current_ids = {
            row.transaction.transaction_id
            for row in rows
            if row.transaction is not None
        }
        for transaction_id, transaction in tuple(self._pending_approval_transactions.items()):
            if transaction_id in current_ids:
                self._record_processing_event(
                    transaction,
                    "목록 유지",
                    "최신 e-Acc 목록에도 이 행이 남아 있습니다. 결재 요청 상태를 e-Acc에서 확인해 주세요.",
                )
            else:
                self._record_processing_event(
                    transaction,
                    "처리 완료",
                    "결재 요청 후 최신 e-Acc 목록에서 이 행이 사라진 것을 확인했습니다.",
                )
            del self._pending_approval_transactions[transaction_id]
        self._apply_filter()

    def _on_close(self) -> None:
        self.browser_service.close()
        self.destroy()


def run(initial_file: str | None = None) -> None:
    app = EAccApplication()
    if initial_file:
        app.after(150, lambda: app._import_file(initial_file))
    app.mainloop()
