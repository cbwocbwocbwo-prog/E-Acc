from __future__ import annotations

import os
import threading
import tkinter as tk
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

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
    NoUnprocessedCardUses,
    NoUnsubmittedTransactions,
    ReceiptNotAvailable,
)
from .employee_directory import EmployeeDirectoryError, load_employee_names, load_mail_recipients
from .mail_notifications import MAIL_SUBJECT, MailRecipient, UnprocessedCardUse, outlook_delivery_status, render_mail_html, send_via_outlook
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
from .smartro_receipt import validate_receipt_with_smartro_support
from .parser import WorkbookFormatError
from .storage import ImportRepository


APP_TITLE = "E-Acc 법인카드 자동처리"
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
        self._account_audit_states: dict[str, tuple[str, str]] = {}
        self._merchant_results: dict[str, MerchantLookupResult] = {}
        self._processing_statuses = self.repository.latest_processing_statuses(
            tuple(APPROVAL_RESULT_STATUSES)
        )
        self._current_target: UnsubmittedTransaction | None = None
        self._current_target_row: ImportDisplayRow | None = None
        self._pending_approval_transactions: dict[str, UnsubmittedTransaction] = {}
        self._processing_result_filter: str | None = None
        self._processing_filter_navigation = False
        self._resume_after_approval_refresh = False
        self._last_approval_removed_from_list = False
        self._session_completed_ids: set[str] = set()
        self._session_exception_ids: set[str] = set()
        self._session_pg_waiting_ids: set[str] = set()
        self._session_target_ids: set[str] = set()
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
        self._mail_recipients: tuple[dict[str, str], ...] = ()
        self._unprocessed_card_uses: tuple[UnprocessedCardUse, ...] = ()
        self._employee_directory_error = ""
        try:
            self._employee_names = load_employee_names()
        except EmployeeDirectoryError as exc:
            # 특근자식비는 확인 가능한 직원 명단이 없으면 정상처리하지 않는다.
            self._employee_directory_error = str(exc)
        try:
            self._mail_recipients = load_mail_recipients()
        except EmployeeDirectoryError as exc:
            self._employee_directory_error = self._employee_directory_error or str(exc)

        self._configure_styles()
        self._build_layout()
        # 비즈노 최초 연결은 느릴 수 있으므로 프로그램 화면을 막지 않고 미리 준비한다.
        self._bizno_client.prewarm()
        self._refresh_history()
        self._refresh_processing_history()
        self._refresh_processing_results()
        self._refresh_mail_log()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("맑은 고딕", 18, "bold"))
        style.configure("Subtitle.TLabel", font=("맑은 고딕", 10), foreground="#5B6470")
        style.configure("CardValue.TLabel", font=("맑은 고딕", 20, "bold"))
        style.configure("CardCaption.TLabel", font=("맑은 고딕", 9), foreground="#606A75")
        style.configure(
            "DashboardFilter.TLabel",
            font=("맑은 고딕", 9, "underline"),
            foreground="#1D5E9E",
        )
        style.configure(
            "GuideLink.TLabel",
            font=("맑은 고딕", 10, "underline"),
            foreground="#1D5E9E",
        )
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
        self.unsubmitted_menu_button = tk.Button(
            sidebar,
            text="미상신 건 수",
            bg="#3D79B8",
            fg="white",
            activebackground="#4B89C9",
            activeforeground="white",
            relief="flat",
            font=("맑은 고딕", 11, "bold"),
            anchor="w",
            padx=20,
            pady=12,
        )
        self.unsubmitted_menu_button.pack(fill="x")
        self.unprocessed_mail_button = tk.Button(
            sidebar,
            text="메일발송(미처리 건 수)",
            bg="#263746",
            fg="#93A3B3",
            disabledforeground="#93A3B3",
            relief="flat",
            command=self._run_unprocessed_mail_process,
            font=("맑은 고딕", 10),
            anchor="w",
            padx=20,
            pady=12,
        )
        self.unprocessed_mail_button.pack(fill="x")
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
        self.current_target_button = ttk.Button(
            button_box,
            text="자동 처리 시작",
            style="Accent.TButton",
            command=self._read_current_first_target,
        )
        self.current_target_button.pack(side="top", fill="x")

        self.current_target_message = tk.StringVar(value="현재 처리대상: 아직 읽지 않음")
        target_box = ttk.LabelFrame(content, text="현재 e-Acc 처리대상", padding=(12, 7))
        target_box.pack(fill="x", pady=(12, 0))
        ttk.Label(target_box, textvariable=self.current_target_message).pack(anchor="w")

        dashboard = ttk.LabelFrame(content, text="이번 실행 현황", padding=(14, 9))
        dashboard.pack(fill="x", pady=(18, 12))
        self.session_vars = {
            "처리대상": tk.StringVar(value="0"),
            "처리 완료": tk.StringVar(value="0"),
            "예외처리": tk.StringVar(value="0"),
            "PG 등록 대기": tk.StringVar(value="0"),
        }
        for index, (caption, variable) in enumerate(self.session_vars.items()):
            card = ttk.Frame(dashboard)
            card.grid(row=0, column=index, padx=(0, 34), sticky="w")
            ttk.Label(card, textvariable=variable, style="CardValue.TLabel").pack(anchor="w")
            result_filter = {
                "처리 완료": "completed",
                "예외처리": "exception",
                "PG 등록 대기": "pg_pending",
            }.get(caption)
            caption_label = ttk.Label(
                card,
                text="처리완료" if caption == "처리 완료" else caption,
                style="DashboardFilter.TLabel" if result_filter else "CardCaption.TLabel",
                cursor="hand2" if result_filter else "",
            )
            caption_label.pack(anchor="w")
            if result_filter:
                caption_label.bind(
                    "<Button-1>",
                    lambda _event, value=result_filter: self._show_processing_result_filter(value),
                )
                caption_label.bind(
                    "<Return>",
                    lambda _event, value=result_filter: self._show_processing_result_filter(value),
                )
        # Current-stage text is the only flexible column.  The guide link is
        # deliberately a fixed final column so it remains visible at the
        # application's minimum window width.
        dashboard.columnconfigure(4, weight=1, minsize=180)
        stage_box = ttk.Frame(dashboard)
        stage_box.grid(row=0, column=4, sticky="ew")
        ttk.Label(stage_box, text="현재 단계", style="CardCaption.TLabel").pack(anchor="w")
        self.current_stage_var = tk.StringVar(value="자동 처리 시작 전")
        ttk.Label(stage_box, textvariable=self.current_stage_var, wraplength=280).pack(anchor="w", pady=(5, 0))
        guide_link = ttk.Label(
            dashboard,
            text="법인카드 검증 기준",
            style="GuideLink.TLabel",
            cursor="hand2",
            takefocus=True,
        )
        guide_link.grid(row=0, column=5, sticky="e", padx=(20, 0))
        guide_link.bind("<Button-1>", self._show_validation_criteria)
        guide_link.bind("<Return>", self._show_validation_criteria)
        guide_link.bind("<space>", self._show_validation_criteria)

        toolbar = ttk.Frame(content)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(
            toolbar,
            text="정상 일반 행은 검증 후 결재요청까지 자동 진행합니다.",
            style="Subtitle.TLabel",
        ).pack(side="left")
        self.actual_merchant_register_button = ttk.Button(
            toolbar,
            text="PG 실구매처 등록",
            command=self._register_selected_actual_merchant,
        )
        self.actual_merchant_register_button.pack(side="right")

        self.progress = ttk.Progressbar(content, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 8))
        self.progress.pack_forget()

        self.notebook = ttk.Notebook(content)
        self.notebook.pack(fill="both", expand=True)
        transactions_tab = ttk.Frame(self.notebook)
        results_tab = ttk.Frame(self.notebook)
        self._processing_results_tab = results_tab
        history_tab = ttk.Frame(self.notebook)
        processing_tab = ttk.Frame(self.notebook)
        mail_log_tab = ttk.Frame(self.notebook)
        tab_specs = (
            (transactions_tab, "실시간_전표처리"),
            (results_tab, "자동_처리결과"),
            (mail_log_tab, "Mail_발송_Log"),
            (processing_tab, "자동_전표처리_Log"),
            (history_tab, "e-Acc DB Export"),
        )
        tab_font = tkfont.nametofont("TkDefaultFont")
        target_width = max(tab_font.measure(label) for _tab, label in tab_specs) + 64
        for tab, label in tab_specs:
            horizontal_padding = max(16, (target_width - tab_font.measure(label)) // 2)
            self.notebook.add(tab, text=label, padding=(horizontal_padding, 8))
        self.notebook.bind("<<NotebookTabChanged>>", self._on_notebook_tab_changed)

        self._build_transactions_table(transactions_tab)
        self._build_processing_results_table(results_tab)
        self._build_history_table(history_tab)
        self._build_processing_table(processing_tab)
        self._build_mail_log_table(mail_log_tab)

        status_frame = ttk.Frame(content)
        status_frame.pack(fill="x", pady=(8, 0))
        self.status_message = tk.StringVar(value="자동 처리 시작을 눌러 e-Acc의 첫 행을 읽어 주세요.")
        self.status_message.trace_add(
            "write", lambda *_args: self.current_stage_var.set(self.status_message.get())
        )
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
            "계정검증",
            "검증근거",
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
            "계정검증": 85,
            "검증근거": 360,
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
        vertical = ttk.Scrollbar(parent, orient="vertical", command=self.processing_tree.yview)
        horizontal = ttk.Scrollbar(parent, orient="horizontal", command=self.processing_tree.xview)
        self.processing_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.processing_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
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

    def _build_mail_log_table(self, parent: ttk.Frame) -> None:
        columns = ("처리시간", "성명", "이메일", "대상 건수", "메일 제목", "상태", "상세 사유")
        self.mail_log_tree = ttk.Treeview(parent, columns=columns, show="headings")
        vertical = ttk.Scrollbar(parent, orient="vertical", command=self.mail_log_tree.yview)
        horizontal = ttk.Scrollbar(parent, orient="horizontal", command=self.mail_log_tree.xview)
        self.mail_log_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.mail_log_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        widths = {"처리시간": 170, "성명": 95, "이메일": 235, "대상 건수": 85, "메일 제목": 260, "상태": 130, "상세 사유": 380}
        for column in columns:
            self.mail_log_tree.heading(column, text=column)
            self.mail_log_tree.column(column, width=widths[column], anchor="center" if column in {"처리시간", "성명", "대상 건수", "상태"} else "w")
        self.mail_log_tree.tag_configure("mail-error", foreground="#A61B1B", background="#FDECEC")
        self.mail_log_tree.tag_configure("mail-warning", foreground="#9A5700", background="#FFF1DA")

    def _build_processing_results_table(self, parent: ttk.Frame) -> None:
        columns = (
            "처리시간",
            "승인번호",
            "증빙일자",
            "사용금액",
            "계정명",
            "거래처",
            "영수증 검증",
            "계정 검증",
            "PG 처리",
            "결재 처리",
            "최종 결과",
            "핵심 사유",
            "내부거래ID",
        )
        self.processing_results_tree = ttk.Treeview(parent, columns=columns, show="headings")
        vertical = ttk.Scrollbar(parent, orient="vertical", command=self.processing_results_tree.yview)
        horizontal = ttk.Scrollbar(parent, orient="horizontal", command=self.processing_results_tree.xview)
        self.processing_results_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.processing_results_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        widths = {
            "처리시간": 175,
            "승인번호": 105,
            "증빙일자": 110,
            "사용금액": 100,
            "계정명": 190,
            "거래처": 210,
            "영수증 검증": 130,
            "계정 검증": 140,
            "PG 처리": 150,
            "결재 처리": 150,
            "최종 결과": 125,
            "핵심 사유": 330,
            "내부거래ID": 205,
        }
        for column in columns:
            self.processing_results_tree.heading(column, text=column)
            self.processing_results_tree.column(
                column,
                width=widths[column],
                anchor="center" if column in {"처리시간", "승인번호", "증빙일자", "사용금액", "영수증 검증", "계정 검증", "PG 처리", "결재 처리", "최종 결과"} else "w",
            )
        # 처리 완료 행은 기본 흰색으로 두고, 사용자의 확인·조치가 필요한
        # 결과만 행 전체의 옅은 배경과 진한 글자색으로 눈에 띄게 한다.
        self.processing_results_tree.tag_configure(
            "결과-예외처리", foreground="#A61B1B", background="#FDECEC"
        )
        self.processing_results_tree.tag_configure(
            "결과-PG등록대기", foreground="#9A5700", background="#FFF1DA"
        )
        self.processing_results_tree.tag_configure(
            "결과-처리대상", foreground="#1D5E9E", background="#ECF5FE"
        )

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

    def _read_current_first_target(self, *, new_run: bool = True) -> None:
        # A user may have closed Electron/Edge after a failed attempt. That
        # leaves a dead CDP endpoint behind, so every fresh user-started run
        # begins with a new browser service. Internal next-row reads keep
        # their existing session (new_run=False).
        if new_run:
            self._reset_browser_service_after_unprocessed_query()
        """Stage 1: get the first row currently visible in e-Acc, without approval."""
        if new_run:
            self._reset_session_dashboard()
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
        self._session_target_ids.add(transaction.transaction_id)
        self._refresh_session_dashboard()
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
                "예외처리",
                "증빙유무가 #이 아니어서 영수증 검증 및 결재 대상에서 제외했습니다.",
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
        if not self._resume_after_approval_refresh:
            return
        self._resume_after_approval_refresh = False
        if not self._last_approval_removed_from_list:
            self.status_message.set(
                "결재요청 행이 최신 목록에 남아 있어 자동 처리를 중지했습니다. e-Acc 상태를 확인해 주세요."
            )
            return
        self.status_message.set("결재 처리 완료를 확인했습니다. 다음 첫 행을 자동 처리합니다...")
        self.after(300, lambda: self._read_current_first_target(new_run=False))

    def _on_browser_error(self, error: Exception) -> None:
        self._set_busy(False)
        if isinstance(error, NoUnsubmittedTransactions):
            self._show_empty_search_result()
            if self._resume_after_approval_refresh:
                self._resume_after_approval_refresh = False
                self.status_message.set("결재 처리 완료를 확인했습니다. 추가 처리 대상이 없습니다.")
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
        if status == "처리 완료":
            self._session_completed_ids.add(transaction.transaction_id)
            self._session_pg_waiting_ids.discard(transaction.transaction_id)
        elif status == "예외처리":
            self._session_exception_ids.add(transaction.transaction_id)
            self._session_pg_waiting_ids.discard(transaction.transaction_id)
        elif status == "PG 등록 대기":
            self._session_pg_waiting_ids.add(transaction.transaction_id)
        elif status == "실구매처 등록 완료":
            self._session_pg_waiting_ids.discard(transaction.transaction_id)
        self._refresh_session_dashboard()
        if status in APPROVAL_RESULT_STATUSES:
            self._processing_statuses[transaction.transaction_id] = status
            # 사용자가 보는 거래내역의 처리결과도 즉시 같은 값으로 갱신한다.
            self._apply_filter()
        self._refresh_processing_history()
        self._refresh_processing_results()
        return event

    def _reset_session_dashboard(self) -> None:
        self._session_completed_ids.clear()
        self._session_exception_ids.clear()
        self._session_pg_waiting_ids.clear()
        self._session_target_ids.clear()
        self._refresh_session_dashboard()

    def _refresh_session_dashboard(self) -> None:
        if not hasattr(self, "session_vars"):
            return
        self.session_vars["처리대상"].set(f"{len(self._session_target_ids):,}")
        self.session_vars["처리 완료"].set(f"{len(self._session_completed_ids):,}")
        self.session_vars["예외처리"].set(f"{len(self._session_exception_ids):,}")
        self.session_vars["PG 등록 대기"].set(f"{len(self._session_pg_waiting_ids):,}")

    def _refresh_after_eacc_approval(
        self,
        transaction: UnsubmittedTransaction | None = None,
    ) -> None:
        """Re-read e-Acc after an approval request to confirm its result."""
        transaction = transaction or self._selected_transaction("처리 결과 재조회")
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
        # 현재 자동 처리 중인 한 행은 검증대기→정상/예외 변화 자체도
        # 감사 이력으로 남긴다. 전체 목록을 단순 조회했을 때 모든 행의
        # 이력이 대량 생성되는 것은 피한다.
        is_current_target = (
            self._current_target is not None
            and self._current_target.transaction_id == transaction.transaction_id
        )
        audit_state = (result.status, result.reason_text)
        if is_current_target and self._account_audit_states.get(transaction.transaction_id) != audit_state:
            self._account_audit_states[transaction.transaction_id] = audit_state
            self._record_processing_event(
                transaction,
                f"계정검증 {result.status}",
                result.reason_text,
            )
        if (
            result.status == "예외"
            and self._processing_statuses.get(transaction.transaction_id) != "예외처리"
        ):
            self._record_processing_event(transaction, "예외처리", result.reason_text)
        return result

    def _start_approval_request(
        self,
        transaction: UnsubmittedTransaction,
        *,
        automatic: bool,
    ) -> None:
        """Send exactly one validated row to the e-Acc approval workflow."""
        blocked_reason = self._approval_request_block_reason(transaction)
        if blocked_reason is not None:
            if automatic:
                self._record_processing_event(transaction, "예외처리", blocked_reason)
                self.status_message.set(f"자동 결재요청 제외: {blocked_reason}")
                self._continue_after_exception(transaction)
                return
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

    def _prepare_selected_approval_request(self) -> None:
        """Allow a user to manually retry approval for the selected current row."""
        transaction = self._selected_transaction("결재요청")
        if transaction is not None:
            self._start_approval_request(transaction, automatic=False)

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
        self._resume_after_approval_refresh = True
        self._last_approval_removed_from_list = False
        self._record_processing_event(transaction, "결재요청 전송", result)
        self.status_message.set("결재요청을 전송했습니다. e-Acc 최신 목록으로 처리 결과를 자동 확인합니다...")
        self.after(700, lambda: self._refresh_after_eacc_approval(transaction))

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
                validation = validate_receipt_with_smartro_support(transaction, result.image_paths)
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
        self._record_processing_event(
            transaction,
            "PG 등록 대기",
            "비즈노 상호조회가 완료되었습니다. e-Acc 실구매처 등록 후 자동 결재를 재개합니다.",
        )
        self._apply_filter()
        self.status_message.set(
            f"PG 상호 조회 완료: 사업자번호 {business_number} / 비즈노 상호 {merchant_name}"
        )
        messagebox.showinfo(
            "PG 실구매처 등록 필요",
            "현재 행은 PG일반 대상입니다.\n\n"
            "e-Acc에서 현재 행의 '선택' 돋보기를 더블클릭해 실구매처 등록 창을 연 뒤,\n"
            "프로그램의 'PG 실구매처 등록' 버튼을 눌러 주세요.\n\n"
            "등록이 성공하면 이 행의 결재요청과 다음 행 자동 처리가 이어집니다.",
            parent=self,
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
                "먼저 '현재 첫 처리대상 읽기'를 실행해 PG 상호 자동조회를 완료해 주세요.",
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
        self.status_message.set(f"실구매처 등록 완료: {message} 결재요청을 자동 진행합니다...")
        self.after(300, lambda: self._start_approval_request(transaction, automatic=True))

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
                validation = validate_receipt_with_smartro_support(transaction, result.image_paths)
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
                self._continue_after_exception(transaction)
            elif validation.status == "정상" and transaction is not None and is_pg_business_type(transaction.business_type):
                self._start_pg_lookup_from_validation(transaction, validation)
            elif validation.status == "정상" and transaction is not None:
                self.status_message.set("영수증·계정 검증 정상: 결재요청을 자동 진행합니다...")
                self.after(300, lambda: self._start_approval_request(transaction, automatic=True))
            elif validation.status != "정상" and transaction is not None:
                mismatch_reasons = " / ".join(
                    check.reason for check in validation.checks if not check.is_match
                )
                self._record_processing_event(
                    transaction,
                    "예외처리",
                    f"영수증 OCR 검증 불일치: {mismatch_reasons or details or validation.status}",
                )
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
                    self._record_processing_event(transaction, "예외처리", reason)
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
                self._record_processing_event(
                    transaction,
                    "예외처리",
                    f"영수증 다운로드 또는 OCR 처리 실패: {error}",
                )
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
                validation = validate_receipt_with_smartro_support(transaction, result.image_paths)
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

    def _run_unprocessed_mail_process(self) -> None:
        """One click: query e-Acc, close Edge, then request Outlook delivery."""
        self._set_busy(True, "e-Acc 미처리내역을 조회한 뒤 Outlook 안내메일을 발송하는 중입니다...")
        self.browser_service.collect_unprocessed_card_uses(
            self._consume_login_credentials(),
            on_success=lambda uses: self.after(0, self._on_unprocessed_card_uses, uses),
            on_error=lambda exc: self.after(0, self._on_unprocessed_card_error, exc),
        )

    def _reset_browser_service_after_unprocessed_query(self) -> None:
        """Close the temporary Edge window and prepare a fresh service for later work."""
        previous = self.browser_service
        try:
            previous.close()
        except Exception:
            # The browser may already have been closed manually.  Its CDP
            # endpoint is then gone, but replacing the service is still safe.
            pass
        finally:
            self.browser_service = EAccountingBrowserService(
                default_database_path().parent / "downloads"
            )

    def _open_unprocessed_card_window(self) -> None:
        existing = getattr(self, "_unprocessed_window", None)
        if existing is not None and existing.winfo_exists():
            existing.deiconify()
            existing.lift()
            return
        window = tk.Toplevel(self)
        self._unprocessed_window = window
        window.title("미처리건수 · 법인카드 안내메일")
        window.geometry("1180x680")
        window.minsize(900, 500)
        content = ttk.Frame(window, padding=16)
        content.pack(fill="both", expand=True)
        ttk.Label(content, text="미처리건수", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            content,
            text="e-Acc 미처리내역을 성명별로 묶어 Outlook 안내메일을 발송합니다. 카드번호와 첨부파일은 사용하지 않습니다.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(3, 12))
        action = ttk.Frame(content)
        action.pack(fill="x", pady=(0, 10))
        self.unprocessed_fetch_button = ttk.Button(action, text="e-Acc 미처리내역 조회", command=self._collect_unprocessed_card_uses)
        self.unprocessed_fetch_button.pack(side="left")
        self.unprocessed_send_button = ttk.Button(action, text="Outlook 메일 발송", command=self._send_unprocessed_mail, state="disabled")
        self.unprocessed_send_button.pack(side="left", padx=8)
        ttk.Button(action, text="발송 결과 확인", command=self._refresh_outlook_delivery_status).pack(side="left")
        self.unprocessed_status_var = tk.StringVar(value="조회 전")
        ttk.Label(action, textvariable=self.unprocessed_status_var, style="Subtitle.TLabel").pack(side="right")
        columns = ("순번", "성명", "부서", "승인번호", "이용일시", "가맹점", "사용금액", "수신자 상태")
        self.unprocessed_tree = ttk.Treeview(content, columns=columns, show="headings")
        vertical = ttk.Scrollbar(content, orient="vertical", command=self.unprocessed_tree.yview)
        horizontal = ttk.Scrollbar(content, orient="horizontal", command=self.unprocessed_tree.xview)
        self.unprocessed_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.unprocessed_tree.grid(row=1, column=0, sticky="nsew")
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal.grid(row=2, column=0, sticky="ew")
        content.rowconfigure(1, weight=1)
        content.columnconfigure(0, weight=1)
        widths = {"순번": 60, "성명": 90, "부서": 170, "승인번호": 105, "이용일시": 145, "가맹점": 250, "사용금액": 105, "수신자 상태": 180}
        for column in columns:
            self.unprocessed_tree.heading(column, text=column)
            self.unprocessed_tree.column(column, width=widths[column], anchor="center" if column not in {"부서", "가맹점", "수신자 상태"} else "w")
        self.unprocessed_tree.tag_configure("recipient-error", foreground="#A61B1B", background="#FDECEC")
        self.unprocessed_tree.tag_configure("recipient-warning", foreground="#9A5700", background="#FFF1DA")

    def _collect_unprocessed_card_uses(self) -> None:
        self.unprocessed_fetch_button.configure(state="disabled")
        self.unprocessed_status_var.set("e-Acc 미처리내역을 조회하는 중입니다...")
        self.browser_service.collect_unprocessed_card_uses(
            self._consume_login_credentials(),
            on_success=lambda uses: self.after(0, self._on_unprocessed_card_uses, uses),
            on_error=lambda exc: self.after(0, self._on_unprocessed_card_error, exc),
        )

    def _on_unprocessed_card_uses(self, uses: tuple[UnprocessedCardUse, ...]) -> None:
        self._unprocessed_card_uses = uses
        # The mail module has no separate screen: after the e-Acc result has
        # arrived, dispose of its temporary Edge session and continue directly
        # to Outlook delivery.  Results remain in the main-window Mail Log.
        self._reset_browser_service_after_unprocessed_query()
        if not uses:
            self._set_busy(False)
            self.status_message.set("e-Acc 미처리내역 대상이 없어 메일을 발송하지 않았습니다.")
            return
        self._send_unprocessed_mail()
        return
        self.unprocessed_fetch_button.configure(state="normal")
        self.unprocessed_send_button.configure(state="normal" if uses else "disabled")
        self.unprocessed_tree.delete(*self.unprocessed_tree.get_children())
        resolvable = 0
        for index, use in enumerate(uses, start=1):
            _recipient, status = self._resolve_mail_recipient(use)
            tag = "" if status == "발송 가능" else "recipient-error"
            if status == "발송 가능":
                resolvable += 1
            self.unprocessed_tree.insert("", "end", values=(index, use.employee_name, use.department, use.approval_number, use.usage_datetime, use.merchant, f"{use.amount:,.0f}원", status), tags=(tag,) if tag else ())
        self.unprocessed_status_var.set(f"총 {len(uses):,}건 / 발송 가능 {resolvable:,}건")

    def _on_unprocessed_card_error(self, error: Exception) -> None:
        self._set_busy(False)
        if isinstance(error, NoUnprocessedCardUses):
            self._unprocessed_card_uses = ()
            self.status_message.set("e-Acc 미처리내역 대상이 없어 메일을 발송하지 않았습니다.")
            return
        messagebox.showerror("미처리내역 조회 실패", str(error), parent=self)
        self.status_message.set(f"미처리내역 조회 실패: {error}")
        return
        if hasattr(self, "unprocessed_fetch_button"):
            self.unprocessed_fetch_button.configure(state="normal")
        if isinstance(error, NoUnprocessedCardUses):
            self._unprocessed_card_uses = ()
            self.unprocessed_send_button.configure(state="disabled")
            self.unprocessed_tree.delete(*self.unprocessed_tree.get_children())
            self.unprocessed_status_var.set("미처리내역 대상 없음")
            return
        messagebox.showerror("미처리내역 조회 실패", str(error), parent=getattr(self, "_unprocessed_window", self))
        self.unprocessed_status_var.set(f"조회 실패: {error}")

    def _resolve_mail_recipient(self, use: UnprocessedCardUse) -> tuple[MailRecipient | None, str]:
        matches = [row for row in self._mail_recipients if row["name"].replace(" ", "") == use.employee_name.replace(" ", "")]
        if not matches:
            return None, "이메일 미확인"
        if len(matches) > 1 and use.department:
            department_matches = [row for row in matches if row["department"] and (row["department"] in use.department or use.department in row["department"])]
            matches = department_matches
        if len(matches) != 1:
            return None, "부서 불일치"
        row = matches[0]
        return MailRecipient(row["name"], row["email"], row["department"]), "발송 가능"

    def _send_unprocessed_mail(self) -> None:
        if not self._unprocessed_card_uses:
            return
        self._set_busy(True, "Outlook 안내메일을 발송하고 메일 Log에 기록하는 중입니다...")
        threading.Thread(target=self._send_unprocessed_mail_worker, daemon=True).start()
        return
        self.unprocessed_send_button.configure(state="disabled")
        self.unprocessed_status_var.set("Outlook 발송 요청을 처리하는 중입니다...")
        threading.Thread(target=self._send_unprocessed_mail_worker, daemon=True).start()

    def _send_unprocessed_mail_worker(self) -> None:
        grouped: dict[MailRecipient, list[UnprocessedCardUse]] = {}
        problems: list[tuple[UnprocessedCardUse, str]] = []
        for use in self._unprocessed_card_uses:
            recipient, status = self._resolve_mail_recipient(use)
            if recipient is None:
                problems.append((use, status))
            else:
                # Past outcomes, including an item still waiting in Outlook's
                # Outbox, never suppress a user-requested reminder.  The
                # button is disabled for this one run, which is sufficient to
                # prevent accidental double-clicks without losing follow-up.
                grouped.setdefault(recipient, []).append(use)
        for use, status in problems:
            self.repository.record_mail_log(recipient_name=use.employee_name, recipient_email="", department=use.department, transaction_ids=(use.transaction_id,), subject=MAIL_SUBJECT, status=status, reason="수신자를 확정할 수 없어 메일을 발송하지 않았습니다.")
        for recipient, uses in grouped.items():
            ids = tuple(use.transaction_id for use in uses)
            try:
                status, outlook_message_id = send_via_outlook(recipient, render_mail_html(recipient.name, uses))
            except Exception as exc:
                self.repository.record_mail_log(recipient_name=recipient.name, recipient_email=recipient.email, department=recipient.department, transaction_ids=ids, subject=MAIL_SUBJECT, status="발송 실패", reason=str(exc))
            else:
                reason = {"발송 완료": "Outlook 보낸 편지함 확인", "발송 대기": "Outlook 보낼 편지함 대기", "발송 확인 불가": "Outlook 폴더에서 발송 결과를 확인하지 못했습니다."}[status]
                self.repository.record_mail_log(recipient_name=recipient.name, recipient_email=recipient.email, department=recipient.department, transaction_ids=ids, subject=MAIL_SUBJECT, status=status, reason=reason, outlook_message_id=outlook_message_id)
        self.after(0, self._on_unprocessed_mail_finished)

    def _on_unprocessed_mail_finished(self) -> None:
        self._refresh_mail_log()
        self._set_busy(False)
        self.status_message.set("Outlook 발송 결과를 메일 Log에 기록했습니다.")
        # The first log entry is intentionally '발송 대기'.  Reconcile once
        # after Outlook has had a moment to move the message into Sent Items.
        self.after(3000, self._refresh_outlook_delivery_status)
        return
        self.unprocessed_send_button.configure(state="normal" if self._unprocessed_card_uses else "disabled")
        self.unprocessed_status_var.set("Outlook 발송 결과를 메일 Log에 기록했습니다.")

    def _refresh_outlook_delivery_status(self) -> None:
        updated = 0
        for item in self.repository.recent_mail_logs():
            if item.status not in {"발송 대기", "발송 확인 불가"} or not item.outlook_message_id:
                continue
            try:
                status = outlook_delivery_status(item.outlook_message_id)
            except Exception:
                # Outlook can briefly reject COM requests while synchronizing.
                # Keep the recorded state and try again on the next mail run.
                continue
            if status != item.status:
                reason = "Outlook 보낸 편지함 확인" if status == "발송 완료" else "Outlook 보낼 편지함 대기" if status == "발송 대기" else "Outlook 폴더에서 발송 결과를 확인하지 못했습니다."
                self.repository.update_mail_status(item.log_id, status, reason)
                updated += 1
        self._refresh_mail_log()
        if hasattr(self, "unprocessed_status_var"):
            self.unprocessed_status_var.set(f"발송 결과 확인 완료: 상태 변경 {updated}건")

    def _refresh_mail_log(self) -> None:
        if not hasattr(self, "mail_log_tree"):
            return
        self.mail_log_tree.delete(*self.mail_log_tree.get_children())
        for item in self.repository.recent_mail_logs():
            tag = "mail-error" if item.status in {"발송 실패", "이메일 미확인", "부서 불일치"} else "mail-warning" if item.status in {"발송 대기", "발송 확인 불가"} else ""
            self.mail_log_tree.insert("", "end", values=(self._format_timestamp(item.sent_at), item.recipient_name, item.recipient_email, item.transaction_count, item.subject, item.status, item.reason), tags=(tag,) if tag else ())

    def _set_busy(self, busy: bool, message: str = "") -> None:
        state = "disabled" if busy else "normal"
        self.actual_merchant_register_button.configure(state=state)
        self.current_target_button.configure(state=state)
        self.unsubmitted_menu_button.configure(state=state)
        self.unprocessed_mail_button.configure(state=state)
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
        visible_count = 0
        for row in rows:
            if row is None:
                continue
            visible_count += 1
            values = list(self._tree_values(row))
            # Excel's physical row number includes its header row.  The UI is
            # a transaction list, so show the first visible transaction as 1.
            values[0] = visible_count
            item_id = self.transactions_tree.insert("", "end", values=values, tags=(row.status,))
            self._rows_by_item[item_id] = row
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
                row.transaction.transaction_id, "처리대상"
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

    def _refresh_history(self) -> None:
        self.history_tree.delete(*self.history_tree.get_children())
        for item in self.repository.recent_history():
            self.history_tree.insert(
                "",
                "end",
                values=(
                    item.job_id,
                    self._format_timestamp(item.imported_at),
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
                    self._format_timestamp(event.event_at),
                    event.status,
                    event.approval_number,
                    event.evidence_date,
                    event.amount,
                    event.merchant,
                    event.reason,
                    event.transaction_id,
                ),
            )

    @staticmethod
    def _format_timestamp(value: str) -> str:
        """Normalize stored ISO timestamps for compact, consistent display."""
        try:
            return datetime.fromisoformat(value).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            return value.replace("T", " ")[:19]

    @staticmethod
    def _compact_reason(reason: str, limit: int = 115) -> str:
        """Keep a result-grid reason useful without duplicating the full Log."""
        compact = reason.split(" 숫자 후보:", maxsplit=1)[0].strip()
        return compact if len(compact) <= limit else compact[: limit - 1].rstrip() + "…"

    def _show_processing_result_filter(self, filter_name: str) -> str:
        """Open the result tab with the dashboard category filter applied."""
        self._processing_result_filter = filter_name
        if self.notebook.select() == str(self._processing_results_tab):
            self._processing_filter_navigation = False
        else:
            self._processing_filter_navigation = True
            self.notebook.select(self._processing_results_tab)
        self._refresh_processing_results()
        return "break"

    def _on_notebook_tab_changed(self, _event: tk.Event[tk.Misc]) -> None:
        """A direct visit to the result tab means show every result again."""
        if self.notebook.select() != str(self._processing_results_tab):
            return
        if self._processing_filter_navigation:
            self._processing_filter_navigation = False
            return
        if self._processing_result_filter is not None:
            self._processing_result_filter = None
            self._refresh_processing_results()

    def _refresh_processing_results(self) -> None:
        """Show one compact, category-based result row for each e-Acc item."""
        if not hasattr(self, "processing_results_tree"):
            return
        self.processing_results_tree.delete(*self.processing_results_tree.get_children())
        grouped: dict[str, list[ProcessingEvent]] = {}
        for event in self.repository.recent_processing_events():
            grouped.setdefault(event.transaction_id, []).append(event)
        for transaction_id, newest_first in grouped.items():
            latest = newest_first[0]
            receipt_event = next(
                (event for event in newest_first if event.status.startswith("영수증 ")),
                None,
            )
            account_event = next(
                (event for event in newest_first if event.status.startswith("계정검증 ")),
                None,
            )
            pg_event = next(
                (
                    event
                    for event in newest_first
                    if event.status.startswith("PG ")
                    or event.status.startswith("실구매처 등록")
                ),
                None,
            )
            approval_event = next(
                (
                    event
                    for event in newest_first
                    if event.status in {
                        "결재선 지정 완료",
                        "결재요청 전송",
                        "결재요청 확인대기",
                        "결재요청 중지",
                        "처리 완료",
                        "목록 유지",
                    }
                ),
                None,
            )
            receipt_status = (
                receipt_event.status.removeprefix("영수증 ")
                if receipt_event is not None
                else "미검증"
            )
            account_status = (
                account_event.status.removeprefix("계정검증 ")
                if account_event is not None
                else "미검증"
            )
            pg_status = (
                pg_event.status.removeprefix("PG ")
                if pg_event is not None and pg_event.status.startswith("PG ")
                else pg_event.status.removeprefix("실구매처 ")
                if pg_event is not None
                else "대상 아님"
            )
            if latest.status == "처리 완료":
                approval_status = "결재 완료"
            elif latest.status == "예외처리" and "다른 사용자" in latest.reason:
                approval_status = "다른 사용자 처리 중"
            elif latest.status == "예외처리" and "결재요청 후" in latest.reason:
                approval_status = "목록 유지"
            elif latest.status == "예외처리":
                approval_status = "결재 미실행"
            elif approval_event is None:
                approval_status = "미실행"
            else:
                approval_status = {
                    "결재선 지정 완료": "결재선 지정",
                    "결재요청 전송": "전송 완료",
                    "결재요청 확인대기": "결과 확인 중",
                    "결재요청 중지": "중지",
                    "처리 완료": "결재 완료",
                    "목록 유지": "목록 유지",
                }[approval_event.status]
            final_status = (
                latest.status
                if latest.status in {"처리 완료", "예외처리", "PG 등록 대기"}
                else "처리중" if approval_event is not None else "처리대상"
            )
            if (
                self._processing_result_filter == "completed"
                and final_status != "처리 완료"
            ):
                continue
            if (
                self._processing_result_filter == "exception"
                and final_status != "예외처리"
            ):
                continue
            if (
                self._processing_result_filter == "pg_pending"
                and pg_status == "대상 아님"
            ):
                continue
            reason = "-" if final_status == "처리 완료" else self._compact_reason(latest.reason)
            account_name = next(
                (event.account_name for event in newest_first if event.account_name),
                "",
            )
            result_tag = {
                "예외처리": "결과-예외처리",
                "PG 등록 대기": "결과-PG등록대기",
                "처리대상": "결과-처리대상",
                "처리중": "결과-처리대상",
            }.get(final_status, "")
            self.processing_results_tree.insert(
                "",
                "end",
                values=(
                    self._format_timestamp(latest.event_at),
                    latest.approval_number,
                    latest.evidence_date,
                    latest.amount,
                    account_name,
                    latest.merchant,
                    receipt_status,
                    account_status,
                    pg_status,
                    approval_status,
                    final_status,
                    reason,
                    transaction_id,
                ),
                tags=(result_tag,) if result_tag else (),
            )

    def _reconcile_pending_approvals(
        self,
        rows: tuple[ImportDisplayRow, ...],
    ) -> None:
        """Record the result of one-row e-Acc approval after a fresh list read."""
        if not self._pending_approval_transactions:
            return
        self._last_approval_removed_from_list = False
        current_ids = {
            row.transaction.transaction_id
            for row in rows
            if row.transaction is not None
        }
        for transaction_id, transaction in tuple(self._pending_approval_transactions.items()):
            if transaction_id in current_ids:
                self._record_processing_event(
                    transaction,
                    "예외처리",
                    "결재요청 후 최신 e-Acc 목록에도 이 행이 남아 있습니다. 결재 요청 상태를 확인해 주세요.",
                )
            else:
                self._record_processing_event(
                    transaction,
                    "처리 완료",
                    "결재 요청 후 최신 e-Acc 목록에서 이 행이 사라진 것을 확인했습니다.",
                )
                self._last_approval_removed_from_list = True
            del self._pending_approval_transactions[transaction_id]
        self._apply_filter()

    def _show_validation_criteria(self, _event: tk.Event[tk.Misc] | None = None) -> str | None:
        """Open a compact, read-only guide to the rules active in this version."""
        existing = getattr(self, "_criteria_window", None)
        if existing is not None and existing.winfo_exists():
            existing.deiconify()
            existing.lift()
            existing.focus_force()
            return "break"

        window = tk.Toplevel(self)
        self._criteria_window = window
        window.title("법인카드 검증 기준")
        window.geometry("1120x455")
        window.minsize(900, 330)
        window.transient(self)

        content = ttk.Frame(window, padding=(18, 16))
        content.pack(fill="both", expand=True)
        ttk.Label(content, text="법인카드 검증 기준", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            content,
            text="현재 프로그램에 적용된 자동 검증 기준입니다. 상세 처리 이력은 Log 탭에서 확인할 수 있습니다.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 12))

        columns = ("구분", "대상 계정·업종", "검증 기준", "예외처리 사유")
        tree = ttk.Treeview(content, columns=columns, show="headings")
        vertical = ttk.Scrollbar(content, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vertical.set)
        tree.pack(side="left", fill="both", expand=True)
        vertical.pack(side="right", fill="y")
        widths = {"구분": 120, "대상 계정·업종": 255, "검증 기준": 480, "예외처리 사유": 230}
        for column in columns:
            tree.heading(column, text=column)
            tree.column(column, width=widths[column], anchor="center" if column == "구분" else "w")
        criteria = (
            ("공통 영수증", "증빙유무 #", "승인번호·증빙일자·사용금액을 영수증 OCR 값과 비교", "영수증 OCR 불일치"),
            ("PG 처리", "업종에 PG일반 포함", "사업자번호 OCR → 비즈노 상호조회 → 실구매처 등록", "PG 조회·등록 확인 필요"),
            ("계정별", "특근자식비", "15,000원당 인정 직원 최소 1명 확인", "특근자식비 사용인원 불충족"),
            ("계정별", "차량유지비-유류대·주차비·세차비·통행료", "영수증 키워드와 계정 유형을 비교", "계정과 상이한 영수증 첨부"),
            ("계정별", "일반복리비-현장지원 (현장대리인 활동지원 식음료대)", "20만원 초과, 주류 문구, 적요 작성 여부 확인", "사용금액 초과 / 주류 포함 / 불필요한 적요 작성"),
            ("계정별", "회의비·부서회의비·업무회의비·일반복리비", "자동 결재 제외", "계정별 예외처리"),
        )
        for row in criteria:
            tree.insert("", "end", values=row)

        footer = ttk.Frame(window, padding=(18, 0, 18, 14))
        footer.pack(fill="x")
        ttk.Button(footer, text="닫기", command=window.destroy).pack(side="right")
        window.protocol("WM_DELETE_WINDOW", window.destroy)
        return "break"

    def _on_close(self) -> None:
        self.browser_service.close()
        self.destroy()


def run(initial_file: str | None = None) -> None:
    app = EAccApplication()
    if initial_file:
        app.after(150, lambda: app._import_file(initial_file))
    app.mainloop()
