# -*- coding: utf-8 -*-
"""
KOCARC 환자 자동등록 — 창(GUI) 프로그램.

명령어 없이 클릭으로 사용:
  - 엑셀 양식 만들기
  - 로그인 정보 입력
  - 시작 / 중지, 실시간 로그 확인 (실제 연구 DB에 생성·저장)

파이썬으로 실행:  uv run python kocarc_gui.py
실행파일(.exe)로 배포하면 파이썬 없이도 더블클릭으로 사용 가능.
"""
import os
import sys
import queue
import threading
import configparser
from urllib.parse import urlsplit

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox, scrolledtext

import kocarc_bot as bot
import build_template

APP_TITLE = "KOCARC AutoRegister"
APP_SUBTITLE = "eCRF 환자 정보 일괄 자동등록"
DEFAULT_LOGIN = "https://ecrf.kr/kocarc/"

# ---------- 색상 팔레트 (세련된 딥 에메랄드 그린 + 세이지 배경, 흰 카드) ----------
BG = "#d3e0d8"        # 창 배경 (세이지 그린 — 흰 카드가 도드라지게 진하게)
CARD = "#ffffff"      # 카드 배경 (흰색)
BORDER = "#c2d2c8"    # 카드/입력칸 테두리 (세이지)
INK = "#1e2a24"       # 본문 글자 (짙은 그린슬레이트)
MUTED = "#5f6f65"     # 보조 글자
FIELD = "#ffffff"     # 입력칸 배경
ACCENT = "#2E6B52"    # 강조(세련된 딥 에메랄드 그린)
ACCENT_DK = "#235340"  # 강조 hover/pressed
GOOD = "#3E8E63"      # 완료 상태(밝은 그린)
GOOD_DK = "#2E6B52"
DANGER = "#c05e5e"    # 오류 상태(부드러운 브릭레드) — 상태 점·경고용
DANGER_DK = "#a64a4a"
STOP = "#7e8c84"      # 중지 버튼(세이지 그레이)
STOP_DK = "#68766e"
SOFT = "#e3ede7"      # 보조 버튼 배경 / 상태바 (연한 세이지)
SOFT_DK = "#d2dfd8"
HEADER = "#ffffff"    # 헤더 (흰 카드처럼 — 세이지 배경 위에서 도드라짐)
HEADER_SUB = "#5f6f65"


def _round_rect_pts(x1, y1, x2, y2, r):
    """둥근 사각형용 폴리곤 점들 (smooth=True 로 그리면 모서리가 둥글어진다)."""
    return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
            x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]


class RoundedButton(tk.Canvas):
    """캔버스로 그린 둥근(알약형) 버튼. ttk 는 모서리를 못 둥글려서 직접 그림.
    ttk 처럼 .configure(state="disabled"/"normal") 로 켜고 끌 수 있다."""

    def __init__(self, parent, text, command=None, *, bg, hover, fg="#ffffff",
                 disabled_bg=None, disabled_fg="#ffffff", container_bg="#ffffff",
                 font=None, padx=20, pady=10, radius=None):
        self._font = tkfont.Font(font=font) if font else tkfont.Font()
        w = self._font.measure(text) + padx * 2
        h = self._font.metrics("linespace") + pady * 2
        r = h // 2 if radius is None else radius          # 기본 = 알약형(완전 둥근 끝)
        super().__init__(parent, width=w, height=h, bg=container_bg,
                         highlightthickness=0, bd=0, takefocus=0)
        self._bg, self._hover, self._fg = bg, hover, fg
        self._dbg = disabled_bg or bg
        self._dfg = disabled_fg
        self._cmd = command
        self._enabled = True
        self._shape = self.create_polygon(
            _round_rect_pts(1, 1, w - 1, h - 1, r),
            smooth=True, splinesteps=24, fill=bg, outline=bg)
        self._label = self.create_text(w / 2, h / 2 + 1, text=text, fill=fg,
                                        font=self._font)
        super().configure(cursor="hand2")
        self.bind("<Enter>", lambda e: self._enabled and self._paint(self._hover))
        self.bind("<Leave>", lambda e: self._enabled and self._paint(self._bg))
        self.bind("<ButtonRelease-1>", self._click)

    def _paint(self, color):
        self.itemconfigure(self._shape, fill=color, outline=color)

    def _click(self, e):
        if (self._enabled and self._cmd
                and 0 <= e.x <= self.winfo_width()
                and 0 <= e.y <= self.winfo_height()):
            self._cmd()

    def _set_enabled(self, on):
        self._enabled = bool(on)
        self._paint(self._bg if on else self._dbg)
        self.itemconfigure(self._label, fill=self._fg if on else self._dfg)
        super().configure(cursor="hand2" if on else "arrow")

    def configure(self, cnf=None, **kw):
        if cnf:
            kw.update(cnf)
        st = kw.pop("state", None)
        if st is not None:
            self._set_enabled(st != "disabled")
        if kw:
            super().configure(**kw)
    config = configure


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        # 고해상도(High-DPI) 화면에서 또렷하게 보이도록 화면 배율 반영
        try:
            dpi = root.winfo_fpixels("1i")          # 화면 DPI (100%=96, 125%=120 …)
            root.tk.call("tk", "scaling", dpi / 72.0)
            self.scale = max(1.0, dpi / 96.0)
        except Exception:
            self.scale = 1.0
        s = self.scale
        root.geometry(f"{int(820 * s)}x{int(740 * s)}")
        root.minsize(int(680 * s), int(460 * s))
        root.configure(bg=BG)
        # 창 아이콘(제목표시줄·작업표시줄) = 로고. 정사각 프레임 .ico 라 찌그러짐 없음.
        try:
            root.iconbitmap(bot.resource_path("assets/logo.ico"))
        except Exception:
            pass
        self.q = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None

        self._setup_fonts()
        self._setup_style()

        # ===== 헤더 (밝은 배경 + 인디고 액센트 바) =====
        header = tk.Frame(root, bg=HEADER)
        header.pack(fill="x")
        tk.Frame(header, bg=BORDER, height=1).pack(fill="x", side="bottom")  # 하단 구분선
        hin = tk.Frame(header, bg=HEADER, padx=22, pady=16)
        hin.pack(fill="x")
        tk.Frame(hin, bg=ACCENT, width=4).pack(side="left", fill="y", padx=(0, 13))
        htxt = tk.Frame(hin, bg=HEADER)
        htxt.pack(side="left", fill="x", expand=True)
        tk.Label(htxt, text=APP_TITLE, bg=HEADER, fg=INK,
                 font=self.f_title, anchor="w").pack(fill="x")
        tk.Label(htxt, text=APP_SUBTITLE, bg=HEADER, fg=HEADER_SUB,
                 font=self.f_small, anchor="w").pack(fill="x", pady=(3, 0))

        # ===== 아래 고정 영역 (먼저 pack 해야 자리를 확보함) =====
        # 상태 표시줄 (맨 아래)
        statusbar = tk.Frame(root, bg=SOFT)
        statusbar.pack(fill="x", side="bottom")
        tk.Frame(statusbar, bg=BORDER, height=1).pack(fill="x", side="top")  # 상단 구분선
        sbin = tk.Frame(statusbar, bg=SOFT, padx=18, pady=8)
        sbin.pack(fill="x")
        self.status_dot = tk.Label(sbin, text="●", bg=SOFT, fg=MUTED, font=self.f_small)
        self.status_dot.pack(side="left")
        self.status_lbl = tk.Label(sbin, text="준비됨", bg=SOFT, fg=INK,
                                   font=self.f_small, anchor="w")
        self.status_lbl.pack(side="left", padx=(6, 0))

        # 로그창 (상태줄 위, 항상 보임)
        logwrap = tk.Frame(root, bg=BG, padx=18)
        logwrap.pack(fill="both", side="bottom", pady=(0, 14))
        logcard = tk.Frame(logwrap, bg=BORDER)
        logcard.pack(fill="both", expand=True)
        loginner = tk.Frame(logcard, bg=CARD)
        loginner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(loginner, text="실시간 로그", bg=CARD, fg=MUTED,
                 font=self.f_small, anchor="w").pack(fill="x", padx=12, pady=(8, 0))
        self.logbox = scrolledtext.ScrolledText(
            loginner, height=10, font=self.f_mono, relief="flat",
            bg="#0f172a", fg="#e2e8f0", insertbackground="#e2e8f0",
            bd=0, padx=12, pady=8)
        self.logbox.pack(fill="both", expand=True, padx=10, pady=10)
        # 로그 색상 태그 (어두운 배경용)
        self.logbox.tag_configure("err", foreground="#f87171")     # 오류·실패·확인필요 (빨강)
        self.logbox.tag_configure("ok", foreground="#34d399")      # 완료·성공 (초록)
        self.logbox.tag_configure("step", foreground="#7dd3fc")    # 환자 시작·PAT_ID (파랑)
        self.logbox.tag_configure("warn", foreground="#fbbf24")    # 검증·건너뜀·중지 (노랑)
        self.logbox.tag_configure("info", foreground="#cbd5e1")    # 일반 (기본)
        self.logbox.configure(state="disabled")

        # 실행 버튼 (로그 위, 항상 보임)
        actions = tk.Frame(root, bg=BG, padx=18)
        actions.pack(fill="x", side="bottom", pady=(8, 10))
        self.start_btn = self._rbtn(actions, "▶  시작", self.start, "primary", BG)
        self.start_btn.pack(side="left")
        self.stop_btn = self._rbtn(actions, "■  중지", self.stop, "stop", BG)
        self.stop_btn.configure(state="disabled")
        self.stop_btn.pack(side="left", padx=(10, 0))
        self._rbtn(actions, "로그 지우기", self.clear_log, "soft", BG).pack(side="right")

        # ===== 설정 카드 (스크롤 가능 영역, 남은 공간 차지) =====
        mid = tk.Frame(root, bg=BG)
        mid.pack(fill="both", expand=True, side="top")
        self.canvas = tk.Canvas(mid, bg=BG, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self._vsb_on = True

        body = tk.Frame(self.canvas, bg=BG, padx=18, pady=16)
        self._body_id = self.canvas.create_window((0, 0), window=body, anchor="nw")
        self.canvas.bind("<Configure>", self._on_canvas_config)
        body.bind("<Configure>", self._on_body_config)
        self.canvas.bind("<MouseWheel>", self._on_wheel)

        # --- 1) 입력 엑셀 ---
        c1 = self._card(body, "1. 입력 엑셀")
        row1 = tk.Frame(c1, bg=CARD)
        row1.pack(fill="x")
        self.excel_var = tk.StringVar(value=self._default_excel())
        ttk.Entry(row1, textvariable=self.excel_var).grid(row=0, column=0, sticky="ew")
        self._rbtn(row1, "찾기…", self.pick_excel, "soft", CARD).grid(
            row=0, column=1, padx=(8, 0))
        self._rbtn(row1, "빈 양식 만들기", self.make_template, "soft", CARD).grid(
            row=0, column=2, padx=(8, 0))
        row1.columnconfigure(0, weight=1)
        tk.Label(c1, text="환자 정보를 채운 엑셀 파일을 선택하세요. 없으면 ‘빈 양식 만들기’로 시작합니다.",
                 bg=CARD, fg=MUTED, font=self.f_small, anchor="w").pack(fill="x", pady=(8, 0))

        # --- 2) 로그인 ---
        c2 = self._card(body, "2. 로그인")
        g2 = tk.Frame(c2, bg=CARD)
        g2.pack(fill="x")
        # 로그인 주소 (한 줄 전체)
        self._field_label(g2, "로그인 주소", 0)
        self.login_var = tk.StringVar(value=DEFAULT_LOGIN)
        ttk.Entry(g2, textvariable=self.login_var).grid(
            row=0, column=1, columnspan=3, sticky="ew", pady=4)
        # 아이디 · 비밀번호 (한 줄에 나란히)
        self._field_label(g2, "아이디", 1)
        saved_id = self._load_saved_id()
        self.id_var = tk.StringVar(value=saved_id)
        ttk.Entry(g2, textvariable=self.id_var).grid(
            row=1, column=1, sticky="ew", padx=(0, 16), pady=4)
        self._field_label(g2, "비밀번호", 1, col=2)
        self.pw_var = tk.StringVar()
        ttk.Entry(g2, textvariable=self.pw_var, show="●").grid(
            row=1, column=3, sticky="ew", pady=4)
        # 아이디 저장
        self.save_id_var = tk.BooleanVar(value=bool(saved_id))
        self._check(g2, "아이디 저장", self.save_id_var).grid(
            row=2, column=1, columnspan=3, sticky="w", pady=(0, 4))
        g2.columnconfigure(1, weight=1)
        g2.columnconfigure(3, weight=1)

        # --- 3) 옵션 ---
        c3 = self._card(body, "3. 옵션")
        g3 = tk.Frame(c3, bg=CARD)
        g3.pack(fill="x")
        self._field_label(g3, "특정 환자키만 (예: 1,2 / 비우면 전체)", 0)
        self.keys_var = tk.StringVar()
        ttk.Entry(g3, textvariable=self.keys_var).grid(row=0, column=1, sticky="ew", pady=4)
        g3.columnconfigure(1, weight=1)

        self.headless_var = tk.BooleanVar(value=False)
        self._check(c3, "브라우저 창 숨기기", self.headless_var
                    ).pack(fill="x", anchor="w", pady=(8, 0))

        self.fresh_var = tk.BooleanVar(value=False)
        self._check(c3, "처음부터 새로 시작 (이전 진행기록 지우기)", self.fresh_var
                    ).pack(fill="x", anchor="w", pady=(10, 0))
        tk.Label(
            c3,
            text=("• 새 환자 명단으로 '처음부터' 등록할 때만 체크하세요.\n"
                  "• 이전 등록 기록(progress.csv)을 지웁니다. 그래야 환자키를 1부터 다시 써도 이미 끝난 환자로 잘못 알고 건너뛰지 않습니다.\n"
                  "• 하던 작업을 '이어서' 할 때는 체크하지 마세요 (기록이 사라집니다).\n"
                  "• 참고: 지난번에 전원 정상 완료됐다면 기록은 이미 자동으로 지워져 있습니다."),
            bg=CARD, fg=MUTED, font=self.f_small, justify="left", anchor="w"
        ).pack(fill="x", anchor="w", padx=(28, 0), pady=(3, 0))

        # 카드 영역 전체에 마우스 휠 스크롤 연결
        self._bind_wheel(body)

        self._append("준비됨. 처음엔 '특정 환자키만'에 1 을 넣어 1명부터 시험하세요.")
        self.root.after(100, self._drain)

    # ---------- 스타일 ----------
    def _setup_fonts(self):
        fam = "Malgun Gothic"                     # Windows 기본, 한글 렌더링 깔끔
        self.f_title = (fam, 16, "bold")
        self.f_section = (fam, 10, "bold")
        self.f_body = (fam, 10)
        self.f_small = (fam, 9)
        self.f_mono = ("Consolas", 9)
        self.f_glyph = ("Segoe UI Symbol", 14)    # 라디오/체크 표시기용 (크게)

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # 입력칸
        style.configure("TEntry", padding=8, relief="flat",
                        fieldbackground=FIELD, foreground=INK,
                        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
        style.map("TEntry",
                  bordercolor=[("focus", ACCENT)],
                  lightcolor=[("focus", ACCENT)],
                  darkcolor=[("focus", ACCENT)])

        # 스크롤바: 화살표 버튼 제거하고 썸(thumb)만 있는 얇은 형태로
        style.layout("Vertical.TScrollbar", [
            ("Vertical.Scrollbar.trough", {
                "sticky": "ns",
                "children": [("Vertical.Scrollbar.thumb",
                              {"expand": "1", "sticky": "nswe"})]})])
        style.configure("Vertical.TScrollbar", troughcolor=BG, background="#a7b6ac",
                        bordercolor=BG, relief="flat", borderwidth=0,
                        width=11, gripcount=0)
        style.map("Vertical.TScrollbar",
                  background=[("active", "#8b9a90"), ("pressed", "#8b9a90")])

        # 라디오 / 체크 (카드 위에 얹힘)
        for name in ("TRadiobutton", "TCheckbutton"):
            style.configure(name, background=CARD, foreground=INK, font=self.f_body)
            style.map(name, background=[("active", CARD)],
                      foreground=[("disabled", "#9aa3b2")])

        # 버튼 공통
        base = dict(font=self.f_body, padding=(16, 9), relief="flat", borderwidth=0)

        # 시작 = 주 색상(인디고)로 통일 — 헤더 액센트와 한 톤
        style.configure("Good.TButton", background=ACCENT, foreground="#ffffff", **base)
        style.map("Good.TButton",
                  background=[("active", ACCENT_DK), ("pressed", ACCENT_DK),
                              ("disabled", "#c7c9ee")],
                  foreground=[("disabled", "#eef0ff")])

        # 중지 = 슬레이트 그레이 (중립적, 파랑 강조색과 안 부딪힘)
        style.configure("Danger.TButton", background=STOP, foreground="#ffffff", **base)
        style.map("Danger.TButton",
                  background=[("active", STOP_DK), ("pressed", STOP_DK),
                              ("disabled", "#c3c9d2")],
                  foreground=[("disabled", "#eef1f5")])

        # 보조 버튼 = 은은한 회색 + 진한 글자
        style.configure("Soft.TButton", background="#e9ecf2", foreground="#334155", **base)
        style.map("Soft.TButton",
                  background=[("active", "#dbe0e9"), ("pressed", "#d1d7e2")])

    def _rbtn(self, parent, text, command, kind, on):
        """둥근 버튼 생성. kind: primary(시작)/stop(중지)/soft(보조). on=놓일 배경색."""
        spec = {
            "primary": (ACCENT, ACCENT_DK, "#ffffff", "#aecdb8", "#eef6f0", 22),
            "stop":    (STOP, STOP_DK, "#ffffff", "#c9d2cb", "#eef2ef", 22),
            "soft":    (SOFT, SOFT_DK, "#33513f", SOFT, "#33513f", 16),
        }[kind]
        bg, hv, fg, dbg, dfg, px = spec
        return RoundedButton(parent, text, command, bg=bg, hover=hv, fg=fg,
                             disabled_bg=dbg, disabled_fg=dfg, container_bg=on,
                             font=self.f_body, padx=px)

    def _card(self, parent, title):
        """흰 카드 + 얇은 테두리. 내부 프레임(bg=CARD)을 돌려준다."""
        outer = tk.Frame(parent, bg=BORDER)
        outer.pack(fill="x", pady=(0, 14))
        inner = tk.Frame(outer, bg=CARD, padx=18, pady=16)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        if title:
            tk.Label(inner, text=title, bg=CARD, fg=ACCENT,
                     font=self.f_section, anchor="w").pack(fill="x", pady=(0, 12))
        return inner

    def _field_label(self, parent, text, row, col=0):
        tk.Label(parent, text=text, bg=CARD, fg=INK, font=self.f_body, anchor="w"
                 ).grid(row=row, column=col, sticky="w", padx=(0, 12), pady=4)

    def _check(self, parent, text, var, command=None):
        """표시기를 직접 그리는 체크박스 (글자 배율에 맞춰 크게 보임)."""
        row = tk.Frame(parent, bg=CARD, cursor="hand2")
        glyph = tk.Label(row, text="", bg=CARD, font=self.f_glyph)
        glyph.pack(side="left")
        lbl = tk.Label(row, text=text, bg=CARD, fg=INK, font=self.f_body)
        lbl.pack(side="left", padx=(8, 0))

        def on_click(_=None):
            var.set(not var.get())
            if command:
                command()

        def refresh(*_):
            on = bool(var.get())
            glyph.configure(text="☑" if on else "☐", fg=ACCENT if on else "#9aa89e")

        for w in (row, glyph, lbl):
            w.bind("<Button-1>", on_click)
        var.trace_add("write", refresh)
        refresh()
        return row

    def _set_status(self, text, kind="idle"):
        colors = {"idle": MUTED, "run": ACCENT, "done": GOOD, "error": DANGER}
        self.status_dot.configure(fg=colors.get(kind, MUTED))
        self.status_lbl.configure(text=text)

    # ---------- 스크롤 ----------
    def _on_canvas_config(self, e):
        # 카드 내용 너비를 캔버스 폭에 맞춤
        self.canvas.itemconfigure(self._body_id, width=e.width)
        self._update_scrollbar()

    def _on_body_config(self, e):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._update_scrollbar()

    def _update_scrollbar(self):
        # 내용이 보이는 영역보다 길 때만 스크롤바 표시
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        content_h = bbox[3] - bbox[1]
        view_h = self.canvas.winfo_height()
        need = content_h > view_h + 1
        if need and not self._vsb_on:
            self.vsb.pack(side="right", fill="y")
            self._vsb_on = True
        elif not need and self._vsb_on:
            self.vsb.pack_forget()
            self._vsb_on = False

    def _on_wheel(self, e):
        if self._vsb_on:
            self.canvas.yview_scroll(int(-e.delta / 120), "units")

    def _bind_wheel(self, widget):
        widget.bind("<MouseWheel>", self._on_wheel)
        for ch in widget.winfo_children():
            self._bind_wheel(ch)

    # ---------- 경로 도우미 ----------
    def _app_dir(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _default_excel(self):
        p = os.path.join(self._app_dir(), "KOCARC_입력양식.xlsx")
        return p if os.path.exists(p) else ""

    # ---------- 아이디 저장(settings.ini, 비밀번호는 저장 안 함) ----------
    def _settings_path(self):
        return os.path.join(self._app_dir(), "settings.ini")

    def _load_saved_id(self):
        try:
            cp = configparser.ConfigParser()
            cp.read(self._settings_path(), encoding="utf-8")
            return cp.get("gui", "member_id", fallback="").strip()
        except Exception:
            return ""

    def _save_id(self, member_id):
        cp = configparser.ConfigParser()
        cp["gui"] = {"member_id": member_id}
        try:
            with open(self._settings_path(), "w", encoding="utf-8") as f:
                cp.write(f)
        except Exception:
            pass

    def _clear_saved_id(self):
        try:
            os.remove(self._settings_path())
        except OSError:
            pass

    # ---------- UI 동작 ----------
    def pick_excel(self):
        p = filedialog.askopenfilename(
            title="입력 엑셀 선택",
            filetypes=[("Excel", "*.xlsx"), ("모든 파일", "*.*")])
        if p:
            self.excel_var.set(p)

    def make_template(self):
        p = filedialog.asksaveasfilename(
            title="빈 양식 저장 위치",
            defaultextension=".xlsx",
            initialfile="KOCARC_입력양식.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        try:
            build_template.build(out_path=p, log=self._append)
            self.excel_var.set(p)
            messagebox.showinfo(APP_TITLE, "빈 양식을 만들었습니다.\n엑셀을 열어 환자 정보를 채우세요.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"양식 생성 실패:\n{e}")

    def clear_log(self):
        self.logbox.configure(state="normal")
        self.logbox.delete("1.0", "end")
        self.logbox.configure(state="disabled")

    # ---------- 실행 ----------
    def _build_conf(self):
        login = self.login_var.get().strip() or DEFAULT_LOGIN
        sp = urlsplit(login)
        base = f"{sp.scheme}://{sp.netloc}" if sp.scheme else "https://ecrf.kr"
        keys = [k.strip() for k in self.keys_var.get().split(",") if k.strip()]
        return {
            "base_url": base.rstrip("/"),
            "login_url": login,
            "member_id": self.id_var.get().strip(),
            "password": self.pw_var.get(),
            "excel": self.excel_var.get().strip(),
            "headless": bool(self.headless_var.get()),
            "fresh_start": bool(self.fresh_var.get()),
            "areas": ["all"],
            "pause": 0.3,
            "only_keys": keys,
            "wait_on_finish": False,
        }

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        conf = self._build_conf()
        if not conf["excel"] or not os.path.exists(conf["excel"]):
            messagebox.showwarning(APP_TITLE, "입력 엑셀 파일을 먼저 선택하세요.")
            return
        if not conf["member_id"]:
            messagebox.showwarning(APP_TITLE, "아이디를 입력하세요.")
            return
        if not conf["password"]:
            messagebox.showwarning(APP_TITLE, "비밀번호를 입력하세요.")
            return
        msg = "실제 연구 DB에 환자가 생성·저장됩니다.\n계속할까요?\n\n"
        if conf.get("fresh_start"):
            msg += ("⚠ [처음부터 새로 시작]이 켜져 있습니다.\n"
                    "이전 진행기록을 모두 지우고 시작합니다.\n"
                    "(이어서 하던 작업이 있으면 사라집니다.)\n\n")
        msg += "(처음에는 '특정 환자키만'에 1 만 넣어 1명으로 시험을 권장)"
        if not messagebox.askyesno(APP_TITLE, msg):
            return

        # 아이디 저장(체크 시) / 해제 시 저장분 삭제 — 비밀번호는 저장 안 함
        if self.save_id_var.get():
            self._save_id(conf["member_id"])
        else:
            self._clear_saved_id()

        # 로그 라우팅
        bot.set_log(lambda line: self.q.put(line))
        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._set_status("실행 중…", "run")

        def work():
            try:
                bot.run_bot(conf, should_stop=self.stop_event.is_set)
            except Exception as e:
                self.q.put(f"[오류] {e}")
            finally:
                self.q.put("__DONE__")

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def stop(self):
        self.stop_event.set()
        self._set_status("중지하는 중…", "error")
        self._append("중지 요청됨 — 현재 환자 처리 후 멈춥니다.")

    # ---------- 로그 폴링 ----------
    def _append(self, text):
        tag = self._classify(text)
        self.logbox.configure(state="normal")
        self.logbox.insert("end", text + "\n", tag)
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    @staticmethod
    def _classify(text):
        """로그 한 줄의 내용에 따라 색상 태그를 정한다."""
        t = text
        if any(k in t for k in ("오류", "[오류]", "실패", "확인필요",
                                "찾을 수 없", "없습니다", "설치되어 있지 않")):
            return "err"
        if any(k in t for k in ("완료", "성공")):
            return "ok"
        if "===" in t or "처리 시작" in t or "PAT_ID =" in t:
            return "step"
        if any(k in t for k in ("검증", "dry", "저장 안 함", "건너뜀", "중지", "종료")):
            return "warn"
        return "info"

    def _drain(self):
        try:
            while True:
                line = self.q.get_nowait()
                if line == "__DONE__":
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self._set_status("완료", "done")
                    self._append("— 종료 —")
                else:
                    self._append(line)
        except queue.Empty:
            pass
        self.root.after(100, self._drain)


def _enable_dpi_awareness():
    """Windows 고해상도 화면에서 화면이 흐릿하게 늘어나 보이는 현상 방지.
    (반드시 Tk 창을 만들기 '전에' 호출해야 함)"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        # 1 = 시스템 DPI 인식: 주 모니터 배율 기준으로 또렷하게 렌더링
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main():
    _enable_dpi_awareness()
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
