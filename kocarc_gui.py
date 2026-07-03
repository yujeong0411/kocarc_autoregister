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
from urllib.parse import urlsplit

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import kocarc_bot as bot
import build_template

APP_TITLE = "KOCARC 환자 자동등록"
APP_SUBTITLE = "eCRF 환자 정보 일괄 자동등록"
DEFAULT_LOGIN = "https://ecrf.kr/kocarc/"
DEFAULT_ID = "kocarc_14"

# ---------- 색상 팔레트 (깔끔한 의료용 블루) ----------
BG = "#eef1f6"        # 창 배경
CARD = "#ffffff"      # 카드 배경
BORDER = "#e2e6ee"    # 카드/입력칸 테두리
INK = "#1f2937"       # 본문 글자
MUTED = "#6b7280"     # 보조 글자
FIELD = "#ffffff"     # 입력칸 배경
ACCENT = "#2563eb"    # 강조(파랑)
ACCENT_DK = "#1d4ed8"
GOOD = "#059669"      # 시작(초록)
GOOD_DK = "#047857"
DANGER = "#dc2626"    # 중지(빨강)
DANGER_DK = "#b91c1c"
SOFT = "#f1f5f9"      # 보조 버튼 배경
SOFT_DK = "#e2e8f0"
HEADER = "#1e3a8a"    # 헤더 배너(짙은 파랑)
HEADER_SUB = "#bfd3f6"


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
            root.iconbitmap(bot.resource_path("logo.ico"))
        except Exception:
            pass
        self.q = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None

        self._setup_fonts()
        self._setup_style()

        # ===== 헤더 배너 =====
        header = tk.Frame(root, bg=HEADER)
        header.pack(fill="x")
        hin = tk.Frame(header, bg=HEADER, padx=22, pady=14)
        hin.pack(fill="x")
        self._logo_img = self._load_logo(int(52 * s))
        if self._logo_img is not None:
            tk.Label(hin, image=self._logo_img, bg=HEADER).pack(side="left", padx=(0, 14))
        htxt = tk.Frame(hin, bg=HEADER)
        htxt.pack(side="left", fill="x", expand=True)
        tk.Label(htxt, text=APP_TITLE, bg=HEADER, fg="#ffffff",
                 font=self.f_title, anchor="w").pack(fill="x")
        tk.Label(htxt, text=APP_SUBTITLE, bg=HEADER, fg=HEADER_SUB,
                 font=self.f_small, anchor="w").pack(fill="x", pady=(2, 0))

        # ===== 아래 고정 영역 (먼저 pack 해야 자리를 확보함) =====
        # 상태 표시줄 (맨 아래)
        statusbar = tk.Frame(root, bg="#dfe4ee")
        statusbar.pack(fill="x", side="bottom")
        sbin = tk.Frame(statusbar, bg="#dfe4ee", padx=18, pady=7)
        sbin.pack(fill="x")
        self.status_dot = tk.Label(sbin, text="●", bg="#dfe4ee", fg=MUTED, font=self.f_small)
        self.status_dot.pack(side="left")
        self.status_lbl = tk.Label(sbin, text="준비됨", bg="#dfe4ee", fg=INK,
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
        self.start_btn = ttk.Button(actions, text="▶  시작", style="Good.TButton", command=self.start)
        self.start_btn.pack(side="left", ipadx=10, ipady=2)
        self.stop_btn = ttk.Button(actions, text="■  중지", style="Danger.TButton",
                                   command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(10, 0), ipadx=10, ipady=2)
        ttk.Button(actions, text="로그 지우기", style="Soft.TButton",
                   command=self.clear_log).pack(side="right")

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
        c1 = self._card(body, "1) 입력 엑셀")
        row1 = tk.Frame(c1, bg=CARD)
        row1.pack(fill="x")
        self.excel_var = tk.StringVar(value=self._default_excel())
        ttk.Entry(row1, textvariable=self.excel_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(row1, text="찾기…", style="Soft.TButton",
                   command=self.pick_excel).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(row1, text="빈 양식 만들기", style="Soft.TButton",
                   command=self.make_template).grid(row=0, column=2, padx=(8, 0))
        row1.columnconfigure(0, weight=1)
        tk.Label(c1, text="환자 정보를 채운 엑셀 파일을 선택하세요. 없으면 ‘빈 양식 만들기’로 시작합니다.",
                 bg=CARD, fg=MUTED, font=self.f_small, anchor="w").pack(fill="x", pady=(8, 0))

        # --- 2) 로그인 ---
        c2 = self._card(body, "2) 로그인")
        g2 = tk.Frame(c2, bg=CARD)
        g2.pack(fill="x")
        self._field_label(g2, "로그인 주소", 0)
        self.login_var = tk.StringVar(value=DEFAULT_LOGIN)
        ttk.Entry(g2, textvariable=self.login_var).grid(row=0, column=1, sticky="ew", pady=4)
        self._field_label(g2, "아이디", 1)
        self.id_var = tk.StringVar(value=DEFAULT_ID)
        ttk.Entry(g2, textvariable=self.id_var).grid(row=1, column=1, sticky="ew", pady=4)
        self._field_label(g2, "비밀번호", 2)
        self.pw_var = tk.StringVar()
        ttk.Entry(g2, textvariable=self.pw_var, show="●").grid(row=2, column=1, sticky="ew", pady=4)
        g2.columnconfigure(1, weight=1)

        # --- 3) 옵션 ---
        c3 = self._card(body, "3) 옵션")
        g3 = tk.Frame(c3, bg=CARD)
        g3.pack(fill="x")
        self._field_label(g3, "특정 환자키만 (예: 1,2 / 비우면 전체)", 0)
        self.keys_var = tk.StringVar()
        ttk.Entry(g3, textvariable=self.keys_var).grid(row=0, column=1, sticky="ew", pady=4)
        g3.columnconfigure(1, weight=1)

        self.headless_var = tk.BooleanVar(value=False)
        self._check(c3, "브라우저 창 숨기기", self.headless_var
                    ).pack(fill="x", anchor="w", pady=(8, 0))

        # 카드 영역 전체에 마우스 휠 스크롤 연결
        self._bind_wheel(body)

        self._append("준비됨. 처음엔 '특정 환자키만'에 1 을 넣어 1명부터 시험하세요.")
        self.root.after(100, self._drain)

    # ---------- 스타일 ----------
    def _setup_fonts(self):
        fam = "Segoe UI"
        self.f_title = (fam, 16, "bold")
        self.f_section = (fam, 11, "bold")
        self.f_body = (fam, 10)
        self.f_small = (fam, 9)
        self.f_mono = ("Consolas", 9)
        self.f_glyph = ("Segoe UI Symbol", 15)   # 라디오/체크 표시기용 (크게)

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # 입력칸
        style.configure("TEntry", padding=7, relief="flat",
                        fieldbackground=FIELD, foreground=INK,
                        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
        style.map("TEntry",
                  bordercolor=[("focus", ACCENT)],
                  lightcolor=[("focus", ACCENT)],
                  darkcolor=[("focus", ACCENT)])

        # 라디오 / 체크 (카드 위에 얹힘)
        for name in ("TRadiobutton", "TCheckbutton"):
            style.configure(name, background=CARD, foreground=INK, font=self.f_body)
            style.map(name, background=[("active", CARD)],
                      foreground=[("disabled", "#9aa3b2")])

        # 버튼 공통
        base = dict(font=self.f_body, padding=(14, 8), relief="flat", borderwidth=0)

        style.configure("Good.TButton", background=GOOD, foreground="#ffffff", **base)
        style.map("Good.TButton",
                  background=[("active", GOOD_DK), ("pressed", GOOD_DK),
                              ("disabled", "#a7d8c6")],
                  foreground=[("disabled", "#eafaf2")])

        style.configure("Danger.TButton", background=DANGER, foreground="#ffffff", **base)
        style.map("Danger.TButton",
                  background=[("active", DANGER_DK), ("pressed", DANGER_DK),
                              ("disabled", "#e6b4b4")],
                  foreground=[("disabled", "#fbeaea")])

        style.configure("Soft.TButton", background=SOFT, foreground=INK, **base)
        style.map("Soft.TButton",
                  background=[("active", SOFT_DK), ("pressed", SOFT_DK)])

    def _card(self, parent, title):
        """흰 카드 + 얇은 테두리. 내부 프레임(bg=CARD)을 돌려준다."""
        outer = tk.Frame(parent, bg=BORDER)
        outer.pack(fill="x", pady=(0, 12))
        inner = tk.Frame(outer, bg=CARD, padx=16, pady=14)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        if title:
            tk.Label(inner, text=title, bg=CARD, fg=ACCENT,
                     font=self.f_section, anchor="w").pack(fill="x", pady=(0, 10))
        return inner

    def _field_label(self, parent, text, row):
        tk.Label(parent, text=text, bg=CARD, fg=INK, font=self.f_body, anchor="w"
                 ).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=4)

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
            glyph.configure(text="☑" if on else "☐", fg=ACCENT if on else "#9aa3b2")

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

    def _load_logo(self, target_h):
        """logo.png 를 PhotoImage 로 로드해 target_h 높이에 맞게 정수배 축소. 없으면 None.
        (bot.resource_path 로 소스·exe(_MEIPASS) 양쪽에서 찾음)"""
        try:
            img = tk.PhotoImage(file=bot.resource_path("logo.png"))
        except Exception:
            return None
        h = img.height()
        if h > target_h:
            img = img.subsample(max(1, round(h / target_h)))
        return img

    # ---------- 경로 도우미 ----------
    def _app_dir(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _default_excel(self):
        p = os.path.join(self._app_dir(), "KOCARC_입력양식.xlsx")
        return p if os.path.exists(p) else ""

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
            "member_id": self.id_var.get().strip() or DEFAULT_ID,
            "password": self.pw_var.get(),
            "excel": self.excel_var.get().strip(),
            "headless": bool(self.headless_var.get()),
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
        if not conf["password"]:
            messagebox.showwarning(APP_TITLE, "비밀번호를 입력하세요.")
            return
        if not messagebox.askyesno(
                APP_TITLE,
                "실제 연구 DB에 환자가 생성·저장됩니다.\n계속할까요?\n\n"
                "(처음에는 '특정 환자키만'에 1 만 넣어 1명으로 시험을 권장)"):
            return

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
