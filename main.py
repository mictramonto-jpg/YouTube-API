#!/usr/bin/env python3
"""
YouTube コメント抽出ツール

YouTube Data API v3 を使用してチャンネル内の動画からコメントを抽出し、
Excel ファイルに出力する GUI アプリケーションです。
"""

import os
import sys
import time
import webbrowser
import threading
import queue
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Error: openpyxl がインストールされていません。")
    print("  pip install -r requirements.txt を実行してください。")
    sys.exit(1)

try:
    from youtube_client import (
        YouTubeClient, APIError, format_date, format_datetime
    )
except ImportError as e:
    print(f"Error: youtube_client.py を読み込めません: {e}")
    sys.exit(1)


# ======================================================================
# Constants
# ======================================================================

APP_TITLE = "YouTube コメント抽出ツール"

HEADERS = [
    "動画URL", "動画タイトル", "動画投稿日", "投稿者名",
    "コメント本文", "コメント高評価数", "コメント投稿日時",
]
COL_IDS = ["url", "title", "date", "author", "comment", "likes", "comment_date"]
COL_WIDTHS_GUI = [220, 260, 100, 130, 340, 110, 140]
COL_WIDTHS_EXCEL = [45, 50, 14, 22, 70, 14, 20]

# カラーパレット
C_BG = "#f5f7fa"
C_PANEL = "#ffffff"
C_ACCENT = "#d32f2f"    # YouTube風の赤
C_ACCENT_DK = "#b71c1c"
C_TEXT = "#212121"
C_MUTED = "#757575"
C_BORDER = "#e0e0e0"
C_OK = "#2e7d32"
C_WARN = "#ef6c00"

API_KEY_GUIDE_URL = "https://console.cloud.google.com/apis/credentials"
API_LIBRARY_URL = "https://console.cloud.google.com/apis/library/youtube.googleapis.com"

API_GUIDE_STEPS = """【YouTube Data API Key 取得手順】

1. Google Cloud Console にアクセス
     → https://console.cloud.google.com/

2. プロジェクトを作成（または既存のプロジェクトを選択）
     画面上部のプロジェクト選択ドロップダウンから
     「新しいプロジェクト」を作成します。

3. YouTube Data API v3 を有効化
     左メニュー「APIとサービス」→「ライブラリ」
     → 検索欄で「YouTube Data API v3」と入力
     → 「有効にする」ボタンをクリック

4. APIキーを作成
     左メニュー「APIとサービス」→「認証情報」
     → 画面上部の「+ 認証情報を作成」→「APIキー」をクリック

5. 作成されたキーをコピー
     ダイアログにAPIキーが表示されます。
     「コピー」ボタンでコピーし、このツールの
     「YouTube Data API Key」欄にペーストしてください。

【注意】
・APIキーは他人に見せないでください
・デフォルトで1日10,000ユニットのクォータ制限があります
・使いすぎると当日は利用できなくなります
"""


# ======================================================================
# Tooltip widget
# ======================================================================

class Tooltip:
    """マウスホバー時に表示されるツールチップ。"""

    def __init__(self, widget, text: str, delay_ms: int = 500):
        self.widget = widget
        self.text = text
        self.delay = delay_ms
        self._after_id = None
        self._tip = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _schedule(self, _evt=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self._tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self._tip, text=self.text, justify="left",
            background="#fffbe6", foreground="#333333",
            relief="solid", borderwidth=1,
            font=("Yu Gothic UI", 9), padx=8, pady=4, wraplength=360,
        ).pack(ipadx=1)

    def _hide(self, _evt=None):
        self._cancel()
        if self._tip:
            self._tip.destroy()
            self._tip = None


# ======================================================================
# Placeholder Entry (hint text)
# ======================================================================

class PlaceholderEntry(ttk.Entry):
    """入力されていないとき、薄いグレーでプレースホルダーを表示するEntry。"""

    def __init__(self, master, placeholder: str = "", color_empty="#9e9e9e",
                 color_filled="#212121", **kwargs):
        super().__init__(master, **kwargs)
        self.placeholder = placeholder
        self.color_empty = color_empty
        self.color_filled = color_filled
        self._has_placeholder = False

        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)

        self._show_placeholder()

    def _show_placeholder(self):
        if not self.get():
            self._has_placeholder = True
            self.configure(foreground=self.color_empty)
            self.insert(0, self.placeholder)

    def _on_focus_in(self, _evt=None):
        if self._has_placeholder:
            self.delete(0, tk.END)
            self.configure(foreground=self.color_filled)
            self._has_placeholder = False

    def _on_focus_out(self, _evt=None):
        if not self.get():
            self._show_placeholder()

    def get_value(self) -> str:
        """プレースホルダーを除いた実際の入力値を返す。"""
        if self._has_placeholder:
            return ""
        return self.get()

    def set_value(self, text: str):
        self.delete(0, tk.END)
        if text:
            self.configure(foreground=self.color_filled)
            self._has_placeholder = False
            self.insert(0, text)
        else:
            self._show_placeholder()


# ======================================================================
# Main Application
# ======================================================================

class YouTubeCommentExtractorApp:
    """YouTube コメント抽出ツールのメインウィンドウ。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.configure(bg=C_BG)

        # DPI/サイズ: 画面サイズに応じてウィンドウサイズを決定
        self._setup_window_size()

        # State
        self.msg_queue = queue.Queue()
        self.running = False
        self.cancelled = False
        self.worker_thread = None
        self.results = []               # List[tuple] - 出力データ
        self._start_time = None
        self._total_videos = 0
        self._processed_videos = 0

        self._setup_styles()
        self._build_menu()
        self._build_header()
        self._build_input_section()
        self._build_results_section()
        self._build_status_bar()

        # 初期モードに合わせてUI (動画数・並び順) の有効/無効を設定
        self._on_mode_changed()

        # キーボードショートカット
        self.root.bind("<F1>", lambda e: self.show_api_guide())
        self.root.bind("<F5>", lambda e: self.start_extraction())
        self.root.bind("<Escape>", lambda e: self.stop_extraction())
        self.root.bind("<Control-s>", lambda e: self.save_to_excel())

        # ウィンドウを閉じる際の処理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._poll_queue()

    # ------------------------------------------------------------------
    # Window / Style setup
    # ------------------------------------------------------------------

    def _setup_window_size(self):
        """画面サイズの約80%でウィンドウサイズを設定。最低720pは確保。"""
        try:
            # Windows DPI awareness
            if sys.platform == "win32":
                from ctypes import windll
                try:
                    windll.shcore.SetProcessDpiAwareness(1)
                except Exception:
                    pass
        except Exception:
            pass

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = min(1400, max(1000, int(sw * 0.78)))
        h = min(850, max(680, int(sh * 0.80)))
        x = (sw - w) // 2
        y = max(0, (sh - h) // 2 - 20)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(960, 620)

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        base_font = ("Yu Gothic UI", 10) if sys.platform == "win32" else ("", 10)
        heading_font = ("Yu Gothic UI", 10, "bold") if sys.platform == "win32" else ("", 10, "bold")

        style.configure("TFrame", background=C_BG)
        style.configure("Panel.TFrame", background=C_PANEL)
        style.configure("TLabel", background=C_BG, foreground=C_TEXT, font=base_font)
        style.configure("Panel.TLabel", background=C_PANEL, foreground=C_TEXT, font=base_font)
        style.configure("Header.TLabel", background=C_PANEL, foreground=C_ACCENT_DK,
                        font=("Yu Gothic UI", 16, "bold") if sys.platform == "win32" else ("", 16, "bold"))
        style.configure("Sub.TLabel", background=C_PANEL, foreground=C_MUTED, font=base_font)
        style.configure("FieldLabel.TLabel", background=C_PANEL, foreground=C_TEXT,
                        font=("Yu Gothic UI", 9, "bold") if sys.platform == "win32" else ("", 9, "bold"))
        style.configure("Status.TLabel", background=C_BG, foreground=C_TEXT, font=base_font)
        style.configure("StatusOK.TLabel", background=C_BG, foreground=C_OK, font=base_font)
        style.configure("StatusWarn.TLabel", background=C_BG, foreground=C_WARN, font=base_font)

        style.configure("TLabelframe", background=C_BG, borderwidth=1, relief="solid")
        style.configure("TLabelframe.Label", background=C_BG, foreground=C_TEXT, font=heading_font)

        style.configure("TEntry", fieldbackground="white", padding=4)

        style.configure("TButton", font=base_font, padding=5)
        style.configure("Run.TButton",
                        font=("Yu Gothic UI", 11, "bold") if sys.platform == "win32" else ("", 11, "bold"),
                        padding=8, foreground="white")
        style.map("Run.TButton",
                  background=[("active", C_ACCENT_DK), ("!disabled", C_ACCENT)],
                  foreground=[("!disabled", "white"), ("disabled", "#e0e0e0")])
        style.configure("Stop.TButton", font=base_font, padding=6)
        style.configure("Link.TButton", font=base_font, padding=4, foreground=C_ACCENT_DK)

        style.configure("Panel.TRadiobutton", background=C_PANEL,
                        foreground=C_TEXT, font=base_font)
        style.map("Panel.TRadiobutton",
                  background=[("active", C_PANEL)])

        style.configure("Results.Treeview",
                        rowheight=28, font=base_font, fieldbackground="white",
                        background="white", foreground=C_TEXT)
        style.configure("Results.Treeview.Heading",
                        font=heading_font, background="#eceff1", foreground=C_TEXT,
                        padding=4)
        style.map("Results.Treeview",
                  background=[("selected", "#bbdefb")],
                  foreground=[("selected", C_TEXT)])

        style.configure("Horizontal.TProgressbar",
                        troughcolor="#e0e0e0", background=C_ACCENT,
                        thickness=18)

    # ------------------------------------------------------------------
    # Menu bar
    # ------------------------------------------------------------------

    def _build_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Excelに保存...\tCtrl+S",
                              command=self.save_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self._on_close)
        menubar.add_cascade(label="ファイル", menu=file_menu)

        run_menu = tk.Menu(menubar, tearoff=0)
        run_menu.add_command(label="稼働 (実行)\tF5", command=self.start_extraction)
        run_menu.add_command(label="停止\tEsc", command=self.stop_extraction)
        menubar.add_cascade(label="実行", menu=run_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="APIキーの取得方法\tF1",
                              command=self.show_api_guide)
        help_menu.add_command(label="Google Cloud Console を開く",
                              command=lambda: webbrowser.open(API_KEY_GUIDE_URL))
        help_menu.add_command(label="YouTube Data API を有効化",
                              command=lambda: webbrowser.open(API_LIBRARY_URL))
        help_menu.add_separator()
        help_menu.add_command(label="使い方", command=self.show_usage)
        help_menu.add_command(label="バージョン情報", command=self.show_about)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)

        self.root.config(menu=menubar)

    # ------------------------------------------------------------------
    # Header
    # ------------------------------------------------------------------

    def _build_header(self):
        hdr = ttk.Frame(self.root, style="Panel.TFrame")
        hdr.pack(fill=tk.X, padx=0, pady=0)

        inner = ttk.Frame(hdr, style="Panel.TFrame", padding=(16, 10))
        inner.pack(fill=tk.X)

        title = ttk.Label(inner, text="▶ " + APP_TITLE, style="Header.TLabel")
        title.pack(side=tk.LEFT)

        sub = ttk.Label(inner, text="YouTube Data API v3 を利用してコメントを抽出 → Excel 出力",
                        style="Sub.TLabel")
        sub.pack(side=tk.LEFT, padx=(14, 0), pady=(6, 0))

        sep = ttk.Separator(self.root, orient="horizontal")
        sep.pack(fill=tk.X)

    # ------------------------------------------------------------------
    # Input section
    # ------------------------------------------------------------------

    def _build_input_section(self):
        outer = ttk.Frame(self.root, padding=(12, 10, 12, 4))
        outer.pack(fill=tk.X)

        # 入力パネル (白背景)
        panel = tk.Frame(outer, bg=C_PANEL, highlightbackground=C_BORDER,
                         highlightthickness=1)
        panel.pack(fill=tk.X)

        inner = ttk.Frame(panel, style="Panel.TFrame", padding=14)
        inner.pack(fill=tk.X)

        # === Row 1: API Key ===
        r1 = ttk.Frame(inner, style="Panel.TFrame")
        r1.pack(fill=tk.X, pady=(0, 10))

        lbl1 = ttk.Label(r1, text="① YouTube Data API Key", style="FieldLabel.TLabel")
        lbl1.pack(side=tk.LEFT)

        guide_btn = ttk.Button(r1, text="❓ 取得方法", style="Link.TButton",
                               command=self.show_api_guide)
        guide_btn.pack(side=tk.LEFT, padx=(10, 0))
        Tooltip(guide_btn, "APIキーの取得手順を表示します (F1)")

        open_console_btn = ttk.Button(r1, text="🌐 Google Cloud Console を開く",
                                       style="Link.TButton",
                                       command=lambda: webbrowser.open(API_KEY_GUIDE_URL))
        open_console_btn.pack(side=tk.LEFT, padx=(6, 0))
        Tooltip(open_console_btn, "ブラウザで Google Cloud Console の認証情報ページを開きます")

        r1b = ttk.Frame(inner, style="Panel.TFrame")
        r1b.pack(fill=tk.X, pady=(0, 12))

        self.api_key_var = tk.StringVar()
        self.api_key_entry = ttk.Entry(r1b, textvariable=self.api_key_var,
                                        show="●", font=("Consolas", 10))
        self.api_key_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        Tooltip(self.api_key_entry,
                "Google Cloud Console で発行したYouTube Data API v3のキーをペースト")

        self._key_visible = False
        self.toggle_key_btn = ttk.Button(r1b, text="👁 表示", width=7,
                                          command=self._toggle_api_key)
        self.toggle_key_btn.pack(side=tk.LEFT, padx=(6, 0))
        Tooltip(self.toggle_key_btn, "APIキーの表示/非表示を切り替えます")

        clear_key_btn = ttk.Button(r1b, text="✕", width=3,
                                    command=lambda: self.api_key_var.set(""))
        clear_key_btn.pack(side=tk.LEFT, padx=(4, 0))
        Tooltip(clear_key_btn, "APIキー欄をクリアします")

        # === Row 2: URL(s) ===
        r2 = ttk.Frame(inner, style="Panel.TFrame")
        r2.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(r2, text="② YouTube URL (チャンネルURL または 動画URL)",
                  style="FieldLabel.TLabel").pack(side=tk.LEFT)

        # ヘルプ (複数URL対応の説明)
        ttk.Label(r2, text="  複数URL対応:「+ URL追加」で枠を追加／CSVで一括読込可",
                  style="Sub.TLabel").pack(side=tk.LEFT, padx=(8, 0))

        # URL 入力行を管理するコンテナ
        self.url_list_frame = ttk.Frame(inner, style="Panel.TFrame")
        self.url_list_frame.pack(fill=tk.X, pady=(2, 4))

        # 複数URLを保持するリスト (各要素は PlaceholderEntry ウィジェット)
        self.url_entries = []
        # 各URL行 Frame を保持 (削除時に参照)
        self.url_rows = []

        # 最初のURL入力行を追加
        self._add_url_row()

        # URL 操作ボタン列 (+ URL追加 / CSVから読込 / 全クリア)
        r2btn = ttk.Frame(inner, style="Panel.TFrame")
        r2btn.pack(fill=tk.X, pady=(0, 12))

        add_url_btn = ttk.Button(r2btn, text="＋ URL追加",
                                 style="Link.TButton",
                                 command=self._add_url_row)
        add_url_btn.pack(side=tk.LEFT)
        Tooltip(add_url_btn,
                "新しいURL入力欄を下に追加します。\n"
                "複数の動画/チャンネルを一括処理できます。")

        csv_btn = ttk.Button(r2btn, text="📁 CSVから一括読込",
                             style="Link.TButton",
                             command=self._load_urls_from_csv)
        csv_btn.pack(side=tk.LEFT, padx=(8, 0))
        Tooltip(csv_btn,
                "CSV/TXTファイルからURLを一括で読み込みます。\n"
                "・1行に1つのURLを記載した形式\n"
                "・CSV形式 (URLを任意の列に記載) も対応\n"
                "・ヘッダ行は自動でスキップ")

        clear_all_btn = ttk.Button(r2btn, text="全クリア",
                                   style="Link.TButton",
                                   command=self._clear_all_urls)
        clear_all_btn.pack(side=tk.LEFT, padx=(8, 0))
        Tooltip(clear_all_btn, "全てのURL入力欄をクリアします (1行は残します)。")

        # 後方互換: 旧コードの self.url_entry 参照を先頭のエントリに紐付け
        self.url_entry = self.url_entries[0]

        # === Row 2c: Extraction Mode (channel-wide vs single-video) ===
        r2c = ttk.Frame(inner, style="Panel.TFrame")
        r2c.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(r2c, text="取得範囲:", style="Panel.TLabel").pack(side=tk.LEFT,
                                                                    padx=(0, 10))

        # "channel" = URL のチャンネル内の複数動画から取得
        # "video"   = URL で指定した動画のみから取得 (デフォルト)
        self.extract_mode_var = tk.StringVar(value="video")

        rb_video = ttk.Radiobutton(
            r2c, text="指定した動画のみ (動画URL) ※複数URL対応",
            variable=self.extract_mode_var, value="video",
            style="Panel.TRadiobutton",
            command=self._on_mode_changed,
        )
        rb_video.pack(side=tk.LEFT, padx=(0, 20))
        Tooltip(rb_video,
                "URLで指定した動画のコメントだけを取得します。\n"
                "動画URL (/watch?v=...、/shorts/...、/live/...) を入力してください。\n"
                "複数のURLを入力すると、それぞれの動画を順番に処理します。")

        rb_channel = ttk.Radiobutton(
            r2c, text="チャンネル全体 (URLからチャンネル特定→複数動画)",
            variable=self.extract_mode_var, value="channel",
            style="Panel.TRadiobutton",
            command=self._on_mode_changed,
        )
        rb_channel.pack(side=tk.LEFT)
        Tooltip(rb_channel,
                "入力URLのチャンネルから「リサーチする動画数」に指定した件数の動画を取得。\n"
                "チャンネルタブURL/動画URLのどちらでもOK (動画URLの場合はそのチャンネル)。\n"
                "複数URLを入力した場合、各チャンネルについて処理します。")

        # === Row 3: Filters (3 inputs) ===
        r3 = ttk.Frame(inner, style="Panel.TFrame")
        r3.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(r3, text="③ 抽出条件", style="FieldLabel.TLabel").pack(side=tk.LEFT)

        r3b = ttk.Frame(inner, style="Panel.TFrame")
        r3b.pack(fill=tk.X)

        # リサーチする動画数
        f1 = ttk.Frame(r3b, style="Panel.TFrame")
        f1.pack(side=tk.LEFT, padx=(0, 24))
        self.max_videos_label = ttk.Label(f1, text="リサーチする動画数",
                                          style="Panel.TLabel")
        self.max_videos_label.pack(anchor="w")
        self.max_videos_var = tk.StringVar(value="30")
        self.e_videos = ttk.Entry(f1, textvariable=self.max_videos_var, width=10,
                                   font=("", 10), justify="right")
        self.e_videos.pack(anchor="w", pady=(2, 0))
        Tooltip(self.e_videos,
                "上位表示されている動画から順に、この件数までを対象にします。\n"
                "例: 30 と入力すると、上位30件の動画のコメントを処理します。\n"
                "URLの動画数が少ない場合は、その件数まで処理します。\n"
                "※「指定した動画のみ」モードでは無効になります。")

        # 並び順 (チャンネル全体モード専用)
        f_order = ttk.Frame(r3b, style="Panel.TFrame")
        f_order.pack(side=tk.LEFT, padx=(0, 24))
        self.order_label = ttk.Label(f_order, text="動画の並び順",
                                     style="Panel.TLabel")
        self.order_label.pack(anchor="w")
        # API値(str) → 表示ラベル(str) の対応表
        self._order_options = [
            ("新しい順", "date"),
            ("視聴回数順 (人気順)", "viewCount"),
            ("評価順", "rating"),
            ("関連度順", "relevance"),
            ("古い順", "oldest"),
        ]
        self._order_label_to_value = {lbl: v for lbl, v in self._order_options}
        self._order_value_to_label = {v: lbl for lbl, v in self._order_options}

        self.order_display_var = tk.StringVar(value=self._order_options[0][0])
        # 内部利用用 (実際のAPI値)
        self.order_var = tk.StringVar(value="date")

        self.order_combo = ttk.Combobox(
            f_order, textvariable=self.order_display_var,
            values=[lbl for lbl, _ in self._order_options],
            state="readonly", width=18, font=("", 10),
        )
        self.order_combo.pack(anchor="w", pady=(2, 0))
        self.order_combo.bind("<<ComboboxSelected>>", self._on_order_changed)
        Tooltip(self.order_combo,
                "チャンネル全体モードでの動画の並び順を指定します:\n"
                "・新しい順 : 最新の動画から (最も高速・低コスト)\n"
                "・視聴回数順: 再生数が多い動画から\n"
                "・評価順  : 高評価が多い動画から\n"
                "・関連度順: YouTube推奨アルゴリズム順\n"
                "・古い順  : 最古の動画から\n"
                "※「指定した動画のみ」モードでは使用されません。\n"
                "※新しい順以外はAPIクォータ消費が多くなります (100 units/リクエスト)。")

        # 高評価数の下限
        f2 = ttk.Frame(r3b, style="Panel.TFrame")
        f2.pack(side=tk.LEFT, padx=(0, 24))
        ttk.Label(f2, text="コメント高評価数の下限", style="Panel.TLabel").pack(anchor="w")
        self.min_likes_var = tk.StringVar(value="0")
        e_likes = ttk.Entry(f2, textvariable=self.min_likes_var, width=10,
                              font=("", 10), justify="right")
        e_likes.pack(anchor="w", pady=(2, 0))
        Tooltip(e_likes,
                "この数値以上の高評価を持つコメントのみを出力します。\n"
                "0 または空欄でフィルタなし (全コメント対象)。")

        # 含まれる文字
        f3 = ttk.Frame(r3b, style="Panel.TFrame")
        f3.pack(side=tk.LEFT, padx=(0, 24))
        ttk.Label(f3, text="コメントに含まれる文字", style="Panel.TLabel").pack(anchor="w")
        self.text_filter_entry = PlaceholderEntry(f3, placeholder="例: : (コロン)",
                                                   width=24, font=("", 10))
        self.text_filter_entry.pack(anchor="w", pady=(2, 0))
        Tooltip(self.text_filter_entry,
                "この文字列を含むコメントのみを出力します。\n"
                "空欄でフィルタなし (全コメント対象)。\n"
                "例: ':' と入力すると、コロンを含むコメント (タイムスタンプ付き等) のみ")

        # === Row 4: Run / Stop buttons ===
        r4 = ttk.Frame(inner, style="Panel.TFrame")
        r4.pack(fill=tk.X, pady=(16, 0))

        ttk.Label(r4, text="ショートカット: F5=稼働 / Esc=停止 / Ctrl+S=保存 / F1=APIキー取得方法",
                  style="Sub.TLabel").pack(side=tk.LEFT)

        self.stop_button = ttk.Button(r4, text="■ 停止", style="Stop.TButton",
                                       command=self.stop_extraction,
                                       state=tk.DISABLED)
        self.stop_button.pack(side=tk.RIGHT, padx=(6, 0))

        self.run_button = ttk.Button(r4, text="▶  稼  働", style="Run.TButton",
                                      command=self.start_extraction)
        self.run_button.pack(side=tk.RIGHT)
        Tooltip(self.run_button, "コメント抽出を開始します (F5)")

    # ------------------------------------------------------------------
    # Results section
    # ------------------------------------------------------------------

    def _build_results_section(self):
        outer = ttk.Frame(self.root, padding=(12, 6, 12, 6))
        outer.pack(fill=tk.BOTH, expand=True)

        # ヘッダー (件数 + 保存ボタン)
        bar = ttk.Frame(outer)
        bar.pack(fill=tk.X, pady=(0, 6))

        lbl = ttk.Label(bar, text="④ 結果", style="Status.TLabel",
                        font=("Yu Gothic UI", 10, "bold") if sys.platform == "win32" else ("", 10, "bold"))
        lbl.pack(side=tk.LEFT)

        self.count_label = ttk.Label(bar, text="  件数: 0", style="Status.TLabel")
        self.count_label.pack(side=tk.LEFT)

        self.export_btn = ttk.Button(bar, text="💾 Excelに名前を付けて保存",
                                      command=self.save_to_excel, state=tk.DISABLED)
        self.export_btn.pack(side=tk.RIGHT)
        Tooltip(self.export_btn, "現在の結果を任意の場所にExcelファイルとして保存します (Ctrl+S)")

        clear_btn = ttk.Button(bar, text="🗑 クリア", command=self._clear_results)
        clear_btn.pack(side=tk.RIGHT, padx=(0, 6))
        Tooltip(clear_btn, "表示されている結果をクリアします")

        # Treeview
        panel = tk.Frame(outer, bg="white", highlightbackground=C_BORDER,
                         highlightthickness=1)
        panel.pack(fill=tk.BOTH, expand=True)

        y_sb = ttk.Scrollbar(panel, orient=tk.VERTICAL)
        x_sb = ttk.Scrollbar(panel, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(
            panel, columns=COL_IDS, show="headings",
            style="Results.Treeview",
            yscrollcommand=y_sb.set, xscrollcommand=x_sb.set,
        )
        y_sb.config(command=self.tree.yview)
        x_sb.config(command=self.tree.xview)

        anchors = {"likes": "center", "date": "center", "comment_date": "center"}
        for cid, heading, width in zip(COL_IDS, HEADERS, COL_WIDTHS_GUI):
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=width, minwidth=50,
                             anchor=anchors.get(cid, "w"))

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_sb.grid(row=0, column=1, sticky="ns")
        x_sb.grid(row=1, column=0, sticky="ew")
        panel.grid_rowconfigure(0, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        # ゼブラストライプ
        self.tree.tag_configure("even", background="#f7fafd")
        self.tree.tag_configure("odd", background="#ffffff")

        # イベント: ダブルクリックで詳細表示
        self.tree.bind("<Double-1>", self._show_row_detail)

        # 右クリックメニュー
        self._build_tree_context_menu()

    def _build_tree_context_menu(self):
        self._tree_menu = tk.Menu(self.root, tearoff=0)
        self._tree_menu.add_command(label="📋 詳細を表示", command=self._show_row_detail)
        self._tree_menu.add_separator()
        self._tree_menu.add_command(label="📋 コメント本文をコピー",
                                     command=lambda: self._copy_cell(4))
        self._tree_menu.add_command(label="📋 動画URLをコピー",
                                     command=lambda: self._copy_cell(0))
        self._tree_menu.add_command(label="📋 動画タイトルをコピー",
                                     command=lambda: self._copy_cell(1))
        self._tree_menu.add_command(label="📋 行全体をコピー (タブ区切り)",
                                     command=self._copy_full_row)
        self._tree_menu.add_separator()
        self._tree_menu.add_command(label="🌐 動画をブラウザで開く",
                                     command=self._open_selected_video)
        self.tree.bind("<Button-3>", self._popup_menu)       # Windows / Linux 右クリック
        self.tree.bind("<Button-2>", self._popup_menu)       # macOS 右クリック

    def _popup_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            try:
                self._tree_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._tree_menu.grab_release()

    def _copy_cell(self, col_index: int):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if 0 <= idx < len(self.results):
            value = str(self.results[idx][col_index])
            self.root.clipboard_clear()
            self.root.clipboard_append(value)

    def _copy_full_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if 0 <= idx < len(self.results):
            row = self.results[idx]
            text = "\t".join(str(v).replace("\t", " ").replace("\n", " ") for v in row)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)

    def _open_selected_video(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if 0 <= idx < len(self.results):
            url = self.results[idx][0]
            if url:
                webbrowser.open(url)

    def _show_row_detail(self, _evt=None):
        """選択行の詳細 (コメント全文) をダイアログで表示。"""
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if not (0 <= idx < len(self.results)):
            return
        row = self.results[idx]

        win = tk.Toplevel(self.root)
        win.title("コメント詳細")
        win.geometry("780x560")
        win.configure(bg=C_BG)
        win.transient(self.root)

        frame = ttk.Frame(win, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Frame(frame)
        info_frame.pack(fill=tk.X, pady=(0, 8))

        infos = [
            ("動画タイトル", row[1]),
            ("動画URL",      row[0]),
            ("動画投稿日",    row[2]),
            ("投稿者",       row[3]),
            ("高評価数",     str(row[5])),
            ("コメント日時",  row[6]),
        ]
        for lbl, val in infos:
            r = ttk.Frame(info_frame)
            r.pack(fill=tk.X, pady=1)
            ttk.Label(r, text=lbl + ":",
                      font=("Yu Gothic UI", 9, "bold") if sys.platform == "win32" else ("", 9, "bold"),
                      width=14, anchor="w").pack(side=tk.LEFT)
            ttk.Label(r, text=val, wraplength=620, justify="left").pack(side=tk.LEFT, anchor="w")

        ttk.Label(frame, text="コメント本文:",
                  font=("Yu Gothic UI", 9, "bold") if sys.platform == "win32" else ("", 9, "bold")).pack(anchor="w", pady=(8, 2))

        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(text_frame, wrap="word", font=("", 10),
                      bg="white", relief="solid", bd=1, padx=8, pady=6)
        sb = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.insert("1.0", row[4])
        txt.configure(state="disabled")
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        def copy_comment():
            self.root.clipboard_clear()
            self.root.clipboard_append(row[4])

        ttk.Button(btn_frame, text="📋 コメントをコピー", command=copy_comment).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="🌐 動画を開く",
                   command=lambda: webbrowser.open(row[0])).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(btn_frame, text="閉じる", command=win.destroy).pack(side=tk.RIGHT)

    def _clear_results(self):
        if self.running:
            return
        if self.results and not messagebox.askyesno(
            "クリア確認", "現在の結果をクリアしますか？\n(Excelに保存されていない場合、データは失われます)"
        ):
            return
        self.tree.delete(*self.tree.get_children())
        self.results.clear()
        self.count_label.config(text="  件数: 0")
        self.export_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.status_var.set("準備完了")

    # ------------------------------------------------------------------
    # Status bar
    # ------------------------------------------------------------------

    def _build_status_bar(self):
        frame = ttk.Frame(self.root, padding=(12, 4, 12, 10))
        frame.pack(fill=tk.X)

        # プログレスバー + 進捗%
        pframe = ttk.Frame(frame)
        pframe.pack(fill=tk.X)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            pframe, variable=self.progress_var, maximum=100,
            style="Horizontal.TProgressbar",
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.pct_label = ttk.Label(pframe, text="0%", style="Status.TLabel", width=6)
        self.pct_label.pack(side=tk.LEFT, padx=(8, 0))

        # ステータステキスト
        sframe = ttk.Frame(frame)
        sframe.pack(fill=tk.X, pady=(6, 0))

        self.status_var = tk.StringVar(value="準備完了 - APIキーとURLを入力して「稼働」をクリックしてください")
        self.status_label = ttk.Label(sframe, textvariable=self.status_var,
                                       style="Status.TLabel")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.elapsed_var = tk.StringVar(value="")
        ttk.Label(sframe, textvariable=self.elapsed_var,
                  style="Status.TLabel").pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _on_mode_changed(self):
        """取得範囲のラジオボタン変更時。動画数・並び順入力の有効/無効を切り替える。"""
        if self.extract_mode_var.get() == "video":
            # 動画のみモード: 動画数・並び順は使わない
            self.e_videos.configure(state=tk.DISABLED)
            self.max_videos_label.configure(foreground=C_MUTED)
            if hasattr(self, "order_combo"):
                self.order_combo.configure(state=tk.DISABLED)
                self.order_label.configure(foreground=C_MUTED)
        else:
            self.e_videos.configure(state=tk.NORMAL)
            self.max_videos_label.configure(foreground=C_TEXT)
            if hasattr(self, "order_combo"):
                self.order_combo.configure(state="readonly")
                self.order_label.configure(foreground=C_TEXT)

    def _on_order_changed(self, _evt=None):
        """並び順コンボボックス変更時。表示ラベルから内部値を更新。"""
        label = self.order_display_var.get()
        value = self._order_label_to_value.get(label, "date")
        self.order_var.set(value)

    # ------------------------------------------------------------------
    # 複数URL対応 (+ URL追加 / CSV一括読込)
    # ------------------------------------------------------------------

    def _add_url_row(self, value: str = ""):
        """URL入力行を1つ追加。value が指定されていれば初期値として設定。"""
        row = ttk.Frame(self.url_list_frame, style="Panel.TFrame")
        row.pack(fill=tk.X, pady=2)

        entry = PlaceholderEntry(
            row,
            placeholder="例: https://www.youtube.com/watch?v=xxxxx",
            font=("", 10),
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if value:
            entry.set_value(value)
        Tooltip(entry,
                "対応形式:\n"
                "・動画URL: https://www.youtube.com/watch?v=xxx\n"
                "・ショート: https://www.youtube.com/shorts/xxx\n"
                "・ライブ: https://www.youtube.com/live/xxx\n"
                "・チャンネル: https://www.youtube.com/@名前/videos")

        # 行番号表示用ラベル (リアルタイム更新)
        num_label = ttk.Label(row, text="", width=4, anchor="e",
                              style="Sub.TLabel")
        num_label.pack(side=tk.LEFT, padx=(4, 0))

        remove_btn = ttk.Button(row, text="✕", width=3,
                                command=lambda r=row, e=entry: self._remove_url_row(r, e))
        remove_btn.pack(side=tk.LEFT, padx=(4, 0))
        Tooltip(remove_btn, "このURL行を削除します")

        self.url_entries.append(entry)
        self.url_rows.append((row, entry, num_label))
        self._renumber_url_rows()

        # フォーカスを新しい行に移動 (初期行を除く)
        if len(self.url_entries) > 1:
            entry.focus_set()

    def _remove_url_row(self, row, entry):
        """指定したURL行を削除 (最低1行は残す)。"""
        if len(self.url_rows) <= 1:
            # 最後の1行は削除せずクリアのみ
            entry.set_value("")
            return

        # リストから削除
        self.url_rows = [r for r in self.url_rows if r[0] is not row]
        self.url_entries = [r[1] for r in self.url_rows]
        row.destroy()
        self._renumber_url_rows()

        # 後方互換性
        if self.url_entries:
            self.url_entry = self.url_entries[0]

    def _renumber_url_rows(self):
        """URL行の通し番号 (#1, #2, ...) を再描画。"""
        for i, (_, _, label) in enumerate(self.url_rows, 1):
            label.configure(text=f"#{i}")

    def _clear_all_urls(self):
        """全てのURL行をクリア (1行は残して中身だけ空に)。"""
        # 最初の1行だけ残し、残りは削除
        while len(self.url_rows) > 1:
            row, _, _ = self.url_rows.pop()
            row.destroy()
        if self.url_rows:
            _, entry, _ = self.url_rows[0]
            entry.set_value("")
        self.url_entries = [r[1] for r in self.url_rows]
        self.url_entry = self.url_entries[0] if self.url_entries else None
        self._renumber_url_rows()

    def _load_urls_from_csv(self):
        """CSV/TXTファイルからURLを一括読込。"""
        filepath = filedialog.askopenfilename(
            title="URL一覧ファイルを選択",
            filetypes=[
                ("CSV / TXT ファイル", "*.csv *.txt"),
                ("CSV ファイル", "*.csv"),
                ("TXT ファイル", "*.txt"),
                ("全てのファイル", "*.*"),
            ],
        )
        if not filepath:
            return

        urls = self._parse_urls_file(filepath)

        if not urls:
            messagebox.showwarning(
                "URLが見つかりません",
                "選択したファイルからYouTubeのURLを検出できませんでした。\n\n"
                "ファイル形式:\n"
                "・1行に1つのURLを記載 (TXT/CSV共通)\n"
                "・CSV形式の場合、任意の列にURL記載可\n"
                "・ヘッダ行は自動でスキップされます"
            )
            return

        # 確認ダイアログ
        proceed = messagebox.askyesno(
            "URL読込の確認",
            f"ファイルから {len(urls)} 件のYouTube URLを検出しました。\n\n"
            f"最初の5件:\n" +
            "\n".join(f"  {i+1}. {u[:70]}" for i, u in enumerate(urls[:5])) +
            (f"\n  ... 他 {len(urls) - 5} 件" if len(urls) > 5 else "") +
            "\n\n既存のURL入力欄を置き換えて読み込みますか？"
        )
        if not proceed:
            return

        # 既存の全行を削除してから読み込んだURLを追加
        for row, _, _ in self.url_rows:
            row.destroy()
        self.url_rows = []
        self.url_entries = []

        for url in urls:
            self._add_url_row(value=url)

        if self.url_entries:
            self.url_entry = self.url_entries[0]

        messagebox.showinfo(
            "読込完了",
            f"{len(urls)} 件のURLを読み込みました。\n\n"
            f"取得範囲モード: "
            f"{'指定した動画のみ' if self.extract_mode_var.get() == 'video' else 'チャンネル全体'}\n"
            "確認したら「▶ 稼働」ボタンを押してください。"
        )

    def _parse_urls_file(self, filepath: str) -> list:
        """
        CSV/TXT からYouTube URLだけを抽出して返す。

        - 各行をカンマで分割し、各セルからURLパターンを抽出
        - youtube.com / youtu.be を含むものだけを対象
        - 重複除去 (順序は維持)
        """
        import re
        yt_pattern = re.compile(
            r"https?://(?:www\.|m\.)?(?:youtube\.com/\S+|youtu\.be/\S+)",
            re.IGNORECASE,
        )

        urls = []
        seen = set()

        # 複数のエンコーディングを試行
        content = None
        for enc in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
            try:
                with open(filepath, "r", encoding=enc) as f:
                    content = f.read()
                break
            except (UnicodeDecodeError, OSError):
                continue

        if content is None:
            messagebox.showerror(
                "読込エラー",
                f"ファイルの読み込みに失敗しました:\n{filepath}\n\n"
                "UTF-8, Shift-JIS のいずれでも読めません。"
            )
            return []

        for line in content.splitlines():
            # カンマ・タブ・セミコロン区切りに対応
            for cell in re.split(r"[,\t;]", line):
                cell = cell.strip().strip('"').strip("'")
                if not cell:
                    continue
                # URL パターン検出
                for m in yt_pattern.finditer(cell):
                    url = m.group(0).rstrip(",;\"'")
                    if url not in seen:
                        seen.add(url)
                        urls.append(url)
                # URL パターンに一致しなくても youtube.com を含むなら追加試行
                if not yt_pattern.search(cell) and (
                    "youtube.com" in cell.lower() or "youtu.be" in cell.lower()
                ):
                    if cell not in seen:
                        seen.add(cell)
                        urls.append(cell)

        return urls

    def _get_all_urls(self) -> list:
        """全URL入力欄から空でないURLを取得。"""
        urls = []
        for entry in self.url_entries:
            u = entry.get_value().strip()
            if u:
                urls.append(u)
        return urls

    def _toggle_api_key(self):
        if self._key_visible:
            self.api_key_entry.config(show="●")
            self.toggle_key_btn.config(text="👁 表示")
        else:
            self.api_key_entry.config(show="")
            self.toggle_key_btn.config(text="🙈 隠す")
        self._key_visible = not self._key_visible

    def _set_running_state(self, running: bool):
        self.running = running
        if running:
            self.run_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.DISABLED)
        else:
            self.run_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            if self.results:
                self.export_btn.config(state=tk.NORMAL)

    def _add_row_to_tree(self, row: tuple):
        """Treeviewへの行追加 (ゼブラストライプ付き、コメントは改行除去)。"""
        idx = len(self.tree.get_children())
        tag = "even" if idx % 2 == 0 else "odd"
        display = list(row)
        display[4] = str(display[4]).replace("\n", " ").replace("\r", " ")[:140]
        if len(str(row[4])) > 140:
            display[4] += " ..."
        self.tree.insert("", tk.END, values=display, tags=(tag,))

    def _update_progress(self, pct: float):
        pct = max(0.0, min(100.0, pct))
        self.progress_var.set(pct)
        self.pct_label.config(text=f"{pct:.0f}%")

    def _update_elapsed(self):
        if self._start_time and self.running:
            elapsed = int(time.time() - self._start_time)
            m, s = divmod(elapsed, 60)
            self.elapsed_var.set(f"経過: {m:02d}:{s:02d}")
            self.root.after(1000, self._update_elapsed)
        elif not self.running:
            pass  # 経過時間はそのまま

    # ------------------------------------------------------------------
    # Queue polling
    # ------------------------------------------------------------------

    def _poll_queue(self):
        try:
            for _ in range(400):
                msg = self.msg_queue.get_nowait()
                kind = msg.get("type")

                if kind == "status":
                    self.status_var.set(msg["text"])
                elif kind == "progress":
                    self._update_progress(msg["value"])
                elif kind == "row":
                    self.results.append(msg["data"])
                    self._add_row_to_tree(msg["data"])
                    self.count_label.config(text=f"  件数: {len(self.results)}")
                elif kind == "done":
                    self._on_done(msg.get("error"))
        except queue.Empty:
            pass
        self.root.after(80, self._poll_queue)

    # ------------------------------------------------------------------
    # Start / Stop
    # ------------------------------------------------------------------

    def start_extraction(self):
        """入力を検証し、バックグラウンドワーカーを起動。"""
        # 多重起動の防止
        if self.running:
            return

        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning(
                "APIキー未入力",
                "YouTube Data API Key を入力してください。\n\n"
                "ヘルプメニュー → 「APIキーの取得方法」から取得手順を確認できます。"
            )
            self.api_key_entry.focus_set()
            return

        urls = self._get_all_urls()
        if not urls:
            messagebox.showwarning(
                "URL未入力",
                "YouTube URL を1つ以上入力してください。\n\n"
                "「+ URL追加」で複数入力、「CSVから一括読込」でまとめて読み込めます。"
            )
            if self.url_entries:
                self.url_entries[0].focus_set()
            return

        mode = self.extract_mode_var.get()  # "channel" or "video"

        # 並び順設定 (チャンネル全体モードのみ有効)
        order = self.order_var.get() if hasattr(self, "order_var") else "date"

        if mode == "channel":
            try:
                max_videos = int(self.max_videos_var.get().strip() or "30")
                if max_videos <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("入力エラー",
                                       "リサーチする動画数は正の整数で入力してください。")
                return
        else:
            # 動画のみモードでは max_videos は使わない
            max_videos = 1

        try:
            min_likes = int(self.min_likes_var.get().strip() or "0")
            if min_likes < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("入力エラー", "高評価数の下限は0以上の整数で入力してください。")
            return

        text_filter = self.text_filter_entry.get_value()

        # 結果をクリア
        self.tree.delete(*self.tree.get_children())
        self.results.clear()
        self.count_label.config(text="  件数: 0")
        self._update_progress(0)
        self.elapsed_var.set("")

        self.cancelled = False
        self._start_time = time.time()
        self._set_running_state(True)
        self._update_elapsed()

        self.worker_thread = threading.Thread(
            target=self._worker,
            args=(api_key, urls, max_videos, min_likes, text_filter, mode, order),
            daemon=True,
        )
        self.worker_thread.start()

    def stop_extraction(self):
        if not self.running:
            return
        self.cancelled = True
        self.status_var.set("停止中...現在の処理完了後に停止します")

    # ------------------------------------------------------------------
    # Background worker
    # ------------------------------------------------------------------

    def _worker(self, api_key, urls, max_videos, min_likes, text_filter,
                mode="video", order="date"):
        put = self.msg_queue.put
        try:
            put({"type": "status", "text": "API接続中..."})
            client = YouTubeClient(api_key)

            put({"type": "status", "text": "APIキーを検証中..."})
            client.validate_api_key()

            # --- 入力URLから動画リストを構築 ---
            all_videos = []
            url_errors = []

            for url_idx, url in enumerate(urls):
                if self.cancelled:
                    break
                put({"type": "status",
                     "text": f"URL解析中... ({url_idx + 1}/{len(urls)}): {url[:60]}"})

                try:
                    url_info = client.parse_url(url)

                    if mode == "video":
                        if url_info["type"] != "video":
                            raise ValueError(
                                f"動画URLではありません (チャンネルURLが入力されています): {url}\n"
                                "「指定した動画のみ」モードでは動画URLを入力してください。"
                            )
                        video = client.get_video_details(url_info["id"])
                        all_videos.append(video)
                    else:
                        put({"type": "status",
                             "text": f"[{url_idx + 1}/{len(urls)}] チャンネル情報取得中..."})
                        channel_id = client.resolve_channel_id(url_info)
                        tab = url_info["tab"]

                        def _vid_progress(msg, n=url_idx + 1, t=len(urls)):
                            put({"type": "status",
                                 "text": f"[{n}/{t}] {msg}"})

                        vids = client.get_videos(
                            channel_id, tab, max_videos,
                            order=order,
                            progress_callback=_vid_progress,
                            cancel_check=lambda: self.cancelled,
                        )
                        all_videos.extend(vids)
                except Exception as exc:
                    url_errors.append(f"{url}: {exc}")
                    put({"type": "status",
                         "text": f"[スキップ] URL{url_idx + 1}: {str(exc)[:80]}"})
                    continue

            if self.cancelled:
                put({"type": "done", "error": None})
                return

            # 重複動画を除去 (video_idで判定、順序は維持)
            seen_vids = set()
            videos = []
            for v in all_videos:
                if v["video_id"] not in seen_vids:
                    seen_vids.add(v["video_id"])
                    videos.append(v)

            total = len(videos)
            if total == 0:
                err_msg = "動画が見つかりませんでした。"
                if url_errors:
                    err_msg += "\n\n【エラー詳細】\n" + "\n".join(url_errors[:5])
                put({"type": "status", "text": "動画が見つかりませんでした。"})
                put({"type": "done",
                     "error": err_msg if url_errors else None})
                return

            self._total_videos = total
            skipped = 0

            for idx, video in enumerate(videos):
                if self.cancelled:
                    break

                pct = (idx / total) * 100
                put({"type": "progress", "value": pct})
                title_short = video["title"][:40]
                put({
                    "type": "status",
                    "text": (
                        f"コメント取得中... 動画 {idx + 1}/{total}"
                        f" ({pct:.0f}%) - {title_short}"
                    ),
                })

                try:
                    comments = client.get_video_comments(
                        video["video_id"],
                        min_likes=min_likes,
                        text_filter=text_filter,
                        cancel_check=lambda: self.cancelled,
                    )
                except Exception as exc:
                    skipped += 1
                    put({
                        "type": "status",
                        "text": (
                            f"[スキップ] 動画 {idx + 1}/{total}: "
                            f"{type(exc).__name__}: {str(exc)[:80]}"
                        ),
                    })
                    continue

                vid_date = format_date(video["published_at"])

                for c in comments:
                    if self.cancelled:
                        break
                    put({
                        "type": "row",
                        "data": (
                            video["url"],
                            video["title"],
                            vid_date,
                            c["author"],
                            c["text"],
                            c["likes"],
                            format_datetime(c["published_at"]),
                        ),
                    })

            put({"type": "progress", "value": 100})
            self._skipped_count = skipped
            put({"type": "done", "error": None})

        except APIError as exc:
            put({"type": "done", "error": f"APIエラー: {exc}"})
        except ValueError as exc:
            put({"type": "done", "error": str(exc)})
        except Exception as exc:
            put({"type": "done", "error": f"予期しないエラー: {type(exc).__name__}: {exc}"})

    # ------------------------------------------------------------------
    # Completion
    # ------------------------------------------------------------------

    def _on_done(self, error):
        self._set_running_state(False)

        if error:
            self.status_var.set(f"エラー: {error}")
            messagebox.showerror(
                "エラー",
                f"処理中にエラーが発生しました:\n\n{error}\n\n"
                "取得済みのデータはテーブルに保持されています。\n"
                "「💾 Excelに名前を付けて保存」から保存できます。",
            )
        elif self.cancelled:
            self.status_var.set(
                f"⏸ 停止しました (取得済み: {len(self.results)} 件)"
            )
        else:
            self.status_var.set(f"✓ 完了！ 合計 {len(self.results)} 件のコメントを取得しました")

        if self.results:
            self._auto_save()

    def _auto_save(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = f"youtube_comments_{ts}.xlsx"
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
        try:
            self._write_excel(path)
            cur = self.status_var.get()
            self.status_var.set(f"{cur}  |  自動保存: {name}")
        except Exception as exc:
            messagebox.showwarning(
                "自動保存エラー",
                f"Excelファイルの自動保存に失敗しました:\n{exc}\n\n"
                "「💾 Excelに名前を付けて保存」から手動で保存してください。"
            )

    # ------------------------------------------------------------------
    # Excel export
    # ------------------------------------------------------------------

    def save_to_excel(self):
        if not self.results:
            messagebox.showinfo("情報", "保存するデータがありません。\n先に「稼働」でデータを取得してください。")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = f"youtube_comments_{ts}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel ファイル", "*.xlsx")],
            initialfile=default,
            title="Excelファイルの保存先を選択",
        )
        if not path:
            return
        try:
            self._write_excel(path)
            messagebox.showinfo("保存完了", f"Excelファイルを保存しました:\n{path}")
        except Exception as exc:
            messagebox.showerror("保存エラー", f"保存に失敗しました:\n{exc}")

    def _write_excel(self, filepath: str):
        wb = Workbook()
        ws = wb.active
        ws.title = "結果"

        label_font = Font(bold=True, size=11, name="Yu Gothic UI")
        thin = Border(
            left=Side("thin", color="BFBFBF"),
            right=Side("thin", color="BFBFBF"),
            top=Side("thin", color="BFBFBF"),
            bottom=Side("thin", color="BFBFBF"),
        )
        hdr_fill = PatternFill("solid", fgColor="4472C4")
        hdr_font = Font(bold=True, size=11, color="FFFFFF", name="Yu Gothic UI")
        data_font = Font(size=10, name="Yu Gothic UI")

        # 設定エリア (エクセルのレイアウトに合わせる)
        ws["A1"] = "YouTube Data API Key"
        ws["A1"].font = label_font

        ws["A4"] = "YouTubeチャンネルURL"
        ws["A4"].font = label_font
        ws.merge_cells("A5:G5")
        ws["A5"] = self.url_entry.get_value()
        ws["A5"].font = data_font

        ws["A7"] = "リサーチする動画数"
        ws["A7"].font = label_font
        try:
            ws["A8"] = int(self.max_videos_var.get())
        except ValueError:
            ws["A8"] = self.max_videos_var.get()

        ws["A10"] = "抽出するコメントの高評価数の下限"
        ws["A10"].font = label_font
        try:
            ws["A11"] = int(self.min_likes_var.get())
        except ValueError:
            ws["A11"] = self.min_likes_var.get()

        ws["A13"] = "抽出するコメントに含まれている文字"
        ws["A13"].font = label_font
        ws["A14"] = self.text_filter_entry.get_value()
        ws["A14"].font = data_font

        # ヘッダー (row 16)
        for ci, text in enumerate(HEADERS, 1):
            cell = ws.cell(row=16, column=ci, value=text)
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # データ (row 17+)
        for ri, row_data in enumerate(self.results, 17):
            for ci, value in enumerate(row_data, 1):
                cell = ws.cell(row=ri, column=ci, value=value)
                cell.border = thin
                cell.font = data_font
                if ci == 5:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                elif ci in (3, 6, 7):
                    cell.alignment = Alignment(horizontal="center", vertical="top")
                else:
                    cell.alignment = Alignment(vertical="top")

        # 列幅
        for i, w in enumerate(COL_WIDTHS_EXCEL, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # ヘッダー行の高さ
        ws.row_dimensions[16].height = 24

        ws.freeze_panes = "A17"
        last_row = 16 + len(self.results)
        ws.auto_filter.ref = f"A16:G{last_row}"

        wb.save(filepath)

    # ------------------------------------------------------------------
    # Dialogs
    # ------------------------------------------------------------------

    def show_api_guide(self):
        win = tk.Toplevel(self.root)
        win.title("APIキーの取得方法")
        win.geometry("720x620")
        win.configure(bg=C_BG)
        win.transient(self.root)

        frame = ttk.Frame(win, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="YouTube Data API Key の取得手順",
                  font=("Yu Gothic UI", 14, "bold") if sys.platform == "win32" else ("", 14, "bold")
                  ).pack(anchor="w", pady=(0, 10))

        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(text_frame, wrap="word", font=("", 10),
                      bg="white", relief="solid", bd=1, padx=12, pady=10)
        sb = ttk.Scrollbar(text_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.insert("1.0", API_GUIDE_STEPS)
        txt.configure(state="disabled")
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        btns = ttk.Frame(frame)
        btns.pack(fill=tk.X, pady=(14, 0))

        ttk.Button(btns, text="🌐 Google Cloud Console を開く",
                   command=lambda: webbrowser.open(API_KEY_GUIDE_URL)
                   ).pack(side=tk.LEFT)
        ttk.Button(btns, text="🌐 YouTube Data API ライブラリを開く",
                   command=lambda: webbrowser.open(API_LIBRARY_URL)
                   ).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="閉じる", command=win.destroy
                   ).pack(side=tk.RIGHT)

    def show_usage(self):
        usage = (
            "【使い方】\n\n"
            "1. ヘルプ → 「APIキーの取得方法」から、YouTube Data API Key を取得します。\n"
            "2. アプリの「① YouTube Data API Key」欄に取得したキーをペーストします。\n"
            "3. 「② YouTube チャンネルURL」欄に、リサーチ対象のURLを入力します。\n"
            "   ・チャンネルの『動画』『ショート』『ライブ』タブのURLに対応\n"
            "   ・動画URLを入力した場合、その動画のチャンネルから動画を取得します\n"
            "4. 「③ 抽出条件」でフィルタを設定します (空欄・0ならフィルタなし)。\n"
            "5. 「▶ 稼働」ボタンをクリックで抽出開始。\n"
            "6. 結果はテーブルにリアルタイムで表示され、完了時にExcelに自動保存されます。\n"
            "   任意の場所に保存したい場合は「💾 Excelに名前を付けて保存」をクリック。\n\n"
            "【便利な機能】\n"
            "・行をダブルクリック→コメント全文を表示\n"
            "・行を右クリック→コピーメニューや動画を開く\n"
            "・F5で稼働、Escで停止、Ctrl+Sで保存\n"
        )
        messagebox.showinfo("使い方", usage)

    def show_about(self):
        about = (
            f"{APP_TITLE}\n\n"
            "YouTube Data API v3 を使用してチャンネル内の動画から\n"
            "コメントを抽出し、Excel に出力するツールです。\n\n"
            "Python: " + sys.version.split()[0] + "\n"
        )
        messagebox.showinfo("バージョン情報", about)

    # ------------------------------------------------------------------
    # Close
    # ------------------------------------------------------------------

    def _on_close(self):
        if self.running:
            if not messagebox.askyesno(
                "終了確認",
                "処理中です。本当に終了しますか？\n(取得済みのデータは保存されません)"
            ):
                return
            self.cancelled = True
        self.root.destroy()


# ======================================================================
# Entry point
# ======================================================================

def main():
    root = tk.Tk()
    YouTubeCommentExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
