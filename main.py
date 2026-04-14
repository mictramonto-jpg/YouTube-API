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
import urllib.request

# Pillow はサムネイル表示専用 (なくても動作する)
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

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

# --------------------------------------------------------------------
# Apple / Notion インスパイアのミニマルパレット
# --------------------------------------------------------------------
C_BG            = "#ffffff"   # メイン背景 (pure white)
C_BG_SUBTLE     = "#fafafa"   # ごく薄いグレー (セクション背景)
C_PANEL         = "#ffffff"   # カード表面
C_PANEL_SUBTLE  = "#f7f7f8"   # Notion風の淡いパネル
C_HOVER         = "#f1f1f2"   # ホバー状態
C_ACCENT        = "#0a0a0a"   # プライマリ (Apple "Continue" 風の黒)
C_ACCENT_HOVER  = "#2a2a2a"
C_ACCENT_SOFT   = "#e8edf3"   # 選択行など
C_LINK          = "#2e6fdb"   # リンク青 (Notion風)
C_LINK_HOVER    = "#1e5bc6"
C_DANGER        = "#d93025"   # エラー/削除
C_DANGER_SOFT   = "#fef3f2"
C_TEXT          = "#1d1d1f"   # 主文字色 (Apple ダークグレー)
C_TEXT_SEC      = "#515154"   # 副文字色
C_MUTED         = "#86868b"   # 補足文字色 (Apple muted)
C_PLACEHOLDER   = "#b0b0b3"
C_BORDER        = "#e5e5e7"   # 境界線 (Apple seperator)
C_BORDER_SOFT   = "#f0f0f1"   # ほぼ見えない境界
C_OK            = "#1d7e40"
C_WARN          = "#c76a00"

# 旧名称の互換エイリアス
C_ACCENT_DK     = C_ACCENT_HOVER
C_PANEL_ALT     = C_BG_SUBTLE
C_BORDER_DK     = C_BORDER
C_ACCENT_LT     = C_ACCENT_SOFT
C_TEXT_SEC      = C_TEXT_SEC  # noqa
C_ACCENT_RED    = C_DANGER

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
        self.root.bind("<Control-s>", lambda e: self.save_results())
        self.root.bind("<Control-h>", lambda e: self.show_history())

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
        """Apple / Notion インスパイアのミニマル tkinter スタイル。"""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # Font stack: SF Pro → Inter → Segoe UI Variable → Yu Gothic UI
        if sys.platform == "darwin":
            _family = "SF Pro Text"
            _family_semi = "SF Pro Display"
        elif sys.platform == "win32":
            _family = "Yu Gothic UI"
            _family_semi = "Yu Gothic UI Semibold"
        else:
            _family = ""
            _family_semi = ""

        def F(size, weight="normal"):
            if weight == "semi":
                return (_family_semi or _family, size) if _family_semi else (_family, size, "bold")
            if weight == "bold":
                return (_family, size, "bold")
            return (_family, size)

        body      = F(10)
        body_sm   = F(9)
        body_xs   = F(8)
        label     = F(10)
        section   = F(11, "semi")
        h1        = F(22, "semi")
        h2        = F(16, "semi")
        h_card    = F(20, "semi")
        button    = F(10)
        button_b  = F(10, "semi")

        # ---------------- フレーム ----------------
        style.configure("TFrame", background=C_BG)
        style.configure("Panel.TFrame", background=C_PANEL)
        style.configure("Subtle.TFrame", background=C_PANEL_SUBTLE)
        style.configure("PanelAlt.TFrame", background=C_BG_SUBTLE)
        style.configure("Card.TFrame", background=C_PANEL)
        style.configure("Divider.TFrame", background=C_BORDER)

        # ---------------- テキスト ----------------
        style.configure("TLabel", background=C_BG, foreground=C_TEXT, font=body)
        style.configure("Panel.TLabel", background=C_PANEL, foreground=C_TEXT, font=body)
        style.configure("Subtle.TLabel", background=C_PANEL_SUBTLE, foreground=C_TEXT, font=body)
        style.configure("PanelAlt.TLabel", background=C_BG_SUBTLE, foreground=C_TEXT, font=body)

        # タイトル系
        style.configure("Header.TLabel", background=C_BG, foreground=C_TEXT, font=h1)
        style.configure("Sub.TLabel", background=C_BG, foreground=C_MUTED, font=body_sm)
        style.configure("HeaderSub.TLabel", background=C_BG, foreground=C_MUTED, font=body)
        style.configure("PanelSub.TLabel", background=C_PANEL, foreground=C_MUTED, font=body_sm)

        # セクション見出し (Notion風 小見出し)
        style.configure("FieldLabel.TLabel", background=C_PANEL, foreground=C_TEXT_SEC,
                        font=F(10, "semi"))
        # プライマリセクションラベル
        style.configure("SectionNum.TLabel", background=C_PANEL, foreground=C_ACCENT, font=section)

        # 状態
        style.configure("Status.TLabel", background=C_BG, foreground=C_TEXT_SEC, font=body_sm)
        style.configure("StatusOK.TLabel", background=C_BG, foreground=C_OK, font=body_sm)
        style.configure("StatusWarn.TLabel", background=C_BG, foreground=C_WARN, font=body_sm)
        style.configure("Danger.TLabel", background=C_PANEL, foreground=C_DANGER, font=body_sm)

        # ---------------- ラベルフレーム ----------------
        style.configure("TLabelframe", background=C_BG, borderwidth=0, relief="flat")
        style.configure("TLabelframe.Label", background=C_BG,
                        foreground=C_TEXT, font=section)

        # ---------------- 入力 (Apple/Notion風の薄い境界) ----------------
        style.configure("TEntry",
                        fieldbackground="white", padding=(10, 8),
                        bordercolor=C_BORDER, lightcolor=C_BORDER,
                        darkcolor=C_BORDER, borderwidth=1, relief="solid")
        style.map("TEntry",
                  bordercolor=[("focus", C_ACCENT)],
                  lightcolor=[("focus", C_ACCENT)],
                  darkcolor=[("focus", C_ACCENT)])
        style.configure("Error.TEntry",
                        fieldbackground=C_DANGER_SOFT,
                        bordercolor=C_DANGER, lightcolor=C_DANGER,
                        darkcolor=C_DANGER, padding=(10, 8))

        # ---------------- ボタン階層 ----------------
        # Secondary (既定)
        style.configure("TButton", font=button, padding=(14, 8),
                        background=C_PANEL, foreground=C_TEXT,
                        bordercolor=C_BORDER, borderwidth=1, relief="flat",
                        focusthickness=0)
        style.map("TButton",
                  background=[("active", C_HOVER), ("disabled", C_PANEL)],
                  foreground=[("disabled", C_MUTED)],
                  bordercolor=[("active", C_BORDER)])

        # Primary (Apple "Continue" 風 - 黒基調)
        style.configure("Primary.TButton",
                        font=button_b, padding=(22, 10),
                        background=C_ACCENT, foreground="white",
                        bordercolor=C_ACCENT, borderwidth=0,
                        relief="flat", focusthickness=0)
        style.map("Primary.TButton",
                  background=[("active", C_ACCENT_HOVER), ("disabled", "#e5e5e7")],
                  foreground=[("disabled", C_MUTED)])

        # 互換: Run = Primary
        style.configure("Run.TButton",
                        font=button_b, padding=(22, 10),
                        background=C_ACCENT, foreground="white",
                        bordercolor=C_ACCENT, borderwidth=0, relief="flat",
                        focusthickness=0)
        style.map("Run.TButton",
                  background=[("active", C_ACCENT_HOVER), ("disabled", "#e5e5e7")],
                  foreground=[("!disabled", "white"), ("disabled", C_MUTED)])

        # Secondary (明示)
        style.configure("Secondary.TButton", font=button, padding=(14, 8),
                        background=C_PANEL, foreground=C_TEXT,
                        bordercolor=C_BORDER, borderwidth=1, relief="flat")
        style.map("Secondary.TButton",
                  background=[("active", C_HOVER)])

        # Danger / Stop
        style.configure("Stop.TButton", font=button, padding=(14, 8),
                        background=C_PANEL, foreground=C_DANGER,
                        bordercolor=C_BORDER, borderwidth=1, relief="flat")
        style.map("Stop.TButton",
                  background=[("active", C_DANGER_SOFT), ("disabled", C_PANEL)],
                  foreground=[("disabled", C_MUTED)])

        # Link (テキストボタン = Notion風)
        style.configure("Link.TButton", font=button, padding=(8, 4),
                        foreground=C_LINK, background=C_PANEL,
                        borderwidth=0, relief="flat", focusthickness=0)
        style.map("Link.TButton",
                  foreground=[("active", C_LINK_HOVER)],
                  background=[("active", C_HOVER)])

        # Toolbar (目立たない操作ボタン)
        style.configure("Toolbar.TButton", font=button_b, padding=(10, 6),
                        foreground=C_TEXT_SEC, background=C_PANEL,
                        borderwidth=0, relief="flat", focusthickness=0)
        style.map("Toolbar.TButton",
                  foreground=[("active", C_TEXT)],
                  background=[("active", C_HOVER)])

        # ---------------- ラジオ / チェック ----------------
        for base_bg, sfx in [(C_PANEL, "Panel"), (C_PANEL_SUBTLE, "Subtle")]:
            style.configure(f"{sfx}.TRadiobutton",
                            background=base_bg, foreground=C_TEXT, font=body,
                            focusthickness=0, indicatorcolor="white",
                            indicatormargin=2)
            style.map(f"{sfx}.TRadiobutton", background=[("active", base_bg)])
            style.configure(f"{sfx}.TCheckbutton",
                            background=base_bg, foreground=C_TEXT, font=body,
                            focusthickness=0)
            style.map(f"{sfx}.TCheckbutton", background=[("active", base_bg)])

        # ---------------- Treeview ----------------
        style.configure("Results.Treeview",
                        rowheight=36, font=body, fieldbackground="white",
                        background="white", foreground=C_TEXT,
                        borderwidth=0, relief="flat")
        style.configure("Results.Treeview.Heading",
                        font=F(9, "semi"), background=C_BG_SUBTLE,
                        foreground=C_MUTED, padding=(10, 10),
                        borderwidth=0, relief="flat")
        style.map("Results.Treeview",
                  background=[("selected", C_ACCENT_SOFT)],
                  foreground=[("selected", C_TEXT)])
        style.map("Results.Treeview.Heading",
                  background=[("active", C_HOVER)])

        # ---------------- プログレスバー (細くミニマル) ----------------
        style.configure("Horizontal.TProgressbar",
                        troughcolor=C_BORDER_SOFT, background=C_ACCENT,
                        thickness=4, borderwidth=0,
                        lightcolor=C_ACCENT, darkcolor=C_ACCENT)

        # ---------------- コンボボックス ----------------
        style.configure("TCombobox", padding=(10, 7), fieldbackground="white",
                        bordercolor=C_BORDER, arrowcolor=C_TEXT_SEC,
                        borderwidth=1, relief="flat")
        style.map("TCombobox",
                  fieldbackground=[("readonly", "white")],
                  selectbackground=[("readonly", "white")],
                  selectforeground=[("readonly", C_TEXT)],
                  bordercolor=[("focus", C_ACCENT)])

        # ---------------- Scrollbar ----------------
        style.configure("Vertical.TScrollbar",
                        background=C_BG_SUBTLE, troughcolor=C_BG,
                        bordercolor=C_BG, arrowcolor=C_MUTED,
                        gripcount=0, relief="flat", borderwidth=0)
        style.configure("Horizontal.TScrollbar",
                        background=C_BG_SUBTLE, troughcolor=C_BG,
                        bordercolor=C_BG, arrowcolor=C_MUTED,
                        gripcount=0, relief="flat", borderwidth=0)

    # ------------------------------------------------------------------
    # Menu bar
    # ------------------------------------------------------------------

    def _build_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="結果を保存...\tCtrl+S",
                              command=self.save_results)
        file_menu.add_separator()
        file_menu.add_command(label="📜 抽出履歴を開く\tCtrl+H",
                              command=self.show_history)
        file_menu.add_command(label="🔑 APIキー管理 (ローテーション)",
                              command=self.show_api_keys_manager)
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
        """Notion / Apple インスパイアのミニマルヘッダー。"""
        hdr = ttk.Frame(self.root, style="TFrame")
        hdr.pack(fill=tk.X, padx=0, pady=0)

        inner = ttk.Frame(hdr, style="TFrame", padding=(28, 22, 28, 16))
        inner.pack(fill=tk.X)

        # タイトル (大きなセンス系のタイポグラフィ、アイコン控えめ)
        title = ttk.Label(inner, text=APP_TITLE, style="Header.TLabel")
        title.pack(side=tk.LEFT)

        # サブタイトル (補足説明は薄いグレーで控えめに)
        sub = ttk.Label(inner,
                        text="YouTubeコメントをシンプルに、美しく抽出。",
                        style="HeaderSub.TLabel")
        sub.pack(side=tk.LEFT, padx=(16, 0), pady=(10, 0))

        # 右上に小さなバージョン/ヒント表示
        hint = ttk.Label(inner, text="⌘F5  /  F5",
                         style="Sub.TLabel")
        hint.pack(side=tk.RIGHT, pady=(12, 0))

        # 薄い区切り線 (Apple風 1px)
        div = tk.Frame(self.root, bg=C_BORDER, height=1)
        div.pack(fill=tk.X)

    # ------------------------------------------------------------------
    # Input section
    # ------------------------------------------------------------------

    def _build_input_section(self):
        """Apple/Notion風: 大きな余白、薄い境界、セクション分離。"""
        outer = ttk.Frame(self.root, padding=(28, 18, 28, 8))
        outer.pack(fill=tk.X)

        # Notion風カード (白背景 + 極薄の境界線)
        panel = tk.Frame(outer, bg=C_PANEL,
                         highlightbackground=C_BORDER,
                         highlightthickness=1, bd=0)
        panel.pack(fill=tk.X)

        inner = ttk.Frame(panel, style="Panel.TFrame", padding=(24, 22))
        inner.pack(fill=tk.X)

        # === Section 1: API Key ===
        lbl1 = ttk.Label(inner, text="API Key",
                        style="FieldLabel.TLabel")
        lbl1.pack(anchor="w", pady=(0, 8))

        r1b = ttk.Frame(inner, style="Panel.TFrame")
        r1b.pack(fill=tk.X, pady=(0, 6))

        self.api_key_var = tk.StringVar()
        self.api_key_entry = ttk.Entry(r1b, textvariable=self.api_key_var,
                                        show="●", font=("Consolas", 10))
        self.api_key_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        Tooltip(self.api_key_entry,
                "Google Cloud Console で発行したYouTube Data API v3のキーをペースト")

        self._key_visible = False
        self.toggle_key_btn = ttk.Button(r1b, text="表示", width=6,
                                          style="Toolbar.TButton",
                                          command=self._toggle_api_key)
        self.toggle_key_btn.pack(side=tk.LEFT, padx=(8, 0))
        Tooltip(self.toggle_key_btn, "APIキーの表示/非表示を切り替えます")

        clear_key_btn = ttk.Button(r1b, text="✕", width=3,
                                    style="Toolbar.TButton",
                                    command=lambda: self.api_key_var.set(""))
        clear_key_btn.pack(side=tk.LEFT, padx=(4, 0))
        Tooltip(clear_key_btn, "APIキー欄をクリアします")

        # ヘルプリンク群 (Notion風に小さく目立たず)
        help_row = ttk.Frame(inner, style="Panel.TFrame")
        help_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(help_row, text="取得方法",
                   style="Link.TButton",
                   command=self.show_api_guide).pack(side=tk.LEFT)
        ttk.Label(help_row, text="·", style="PanelSub.TLabel").pack(
            side=tk.LEFT, padx=4)
        ttk.Button(help_row, text="Google Cloud Console",
                   style="Link.TButton",
                   command=lambda: webbrowser.open(API_KEY_GUIDE_URL)).pack(side=tk.LEFT)

        # 区切り (極薄)
        tk.Frame(inner, bg=C_BORDER_SOFT, height=1).pack(
            fill=tk.X, pady=(22, 0))

        # === Section 2: URL ===
        ttk.Label(inner, text="URL",
                  style="FieldLabel.TLabel").pack(anchor="w", pady=(18, 4))
        ttk.Label(inner,
                  text="チャンネルURL または 動画URL。複数入力・CSV一括読込可。",
                  style="PanelSub.TLabel").pack(anchor="w", pady=(0, 8))

        # URL 入力行を管理するコンテナ
        self.url_list_frame = ttk.Frame(inner, style="Panel.TFrame")
        self.url_list_frame.pack(fill=tk.X, pady=(0, 4))

        # 複数URLを保持するリスト (各要素は PlaceholderEntry ウィジェット)
        self.url_entries = []
        # 各URL行 Frame を保持 (削除時に参照)
        self.url_rows = []

        # 最初のURL入力行を追加
        self._add_url_row()

        # URL 操作ボタン列
        r2btn = ttk.Frame(inner, style="Panel.TFrame")
        r2btn.pack(fill=tk.X, pady=(4, 0))

        add_url_btn = ttk.Button(r2btn, text="＋ URLを追加",
                                 style="Link.TButton",
                                 command=self._add_url_row)
        add_url_btn.pack(side=tk.LEFT)
        Tooltip(add_url_btn, "新しいURL入力欄を追加します。")

        ttk.Label(r2btn, text="·", style="PanelSub.TLabel").pack(side=tk.LEFT, padx=4)

        csv_btn = ttk.Button(r2btn, text="CSVから一括読込",
                             style="Link.TButton",
                             command=self._load_urls_from_csv)
        csv_btn.pack(side=tk.LEFT)
        Tooltip(csv_btn,
                "CSV/TXTファイルからURLを一括で読み込みます。\n"
                "1行に1つのURL、またはCSVの任意列に記載。")

        ttk.Label(r2btn, text="·", style="PanelSub.TLabel").pack(side=tk.LEFT, padx=4)

        clear_all_btn = ttk.Button(r2btn, text="すべてクリア",
                                   style="Link.TButton",
                                   command=self._clear_all_urls)
        clear_all_btn.pack(side=tk.LEFT)
        Tooltip(clear_all_btn, "全てのURL入力欄をクリアします。")

        # 後方互換: 旧コードの self.url_entry 参照を先頭のエントリに紐付け
        self.url_entry = self.url_entries[0]

        # 区切り
        tk.Frame(inner, bg=C_BORDER_SOFT, height=1).pack(
            fill=tk.X, pady=(22, 0))

        # === Section 3: 取得範囲 ===
        ttk.Label(inner, text="取得範囲",
                  style="FieldLabel.TLabel").pack(anchor="w", pady=(18, 8))

        r2c = ttk.Frame(inner, style="Panel.TFrame")
        r2c.pack(fill=tk.X, pady=(0, 12))

        # "channel" = URL のチャンネル内の複数動画から取得
        # "video"   = URL で指定した動画のみから取得 (デフォルト)
        self.extract_mode_var = tk.StringVar(value="video")

        rb_video = ttk.Radiobutton(
            r2c, text="指定した動画のみ",
            variable=self.extract_mode_var, value="video",
            style="Panel.TRadiobutton",
            command=self._on_mode_changed,
        )
        rb_video.pack(side=tk.LEFT, padx=(0, 28))
        Tooltip(rb_video,
                "URLで指定した動画のコメントだけを取得します。\n"
                "動画URL (/watch?v=...、/shorts/...、/live/...) を入力してください。")

        rb_channel = ttk.Radiobutton(
            r2c, text="チャンネル全体",
            variable=self.extract_mode_var, value="channel",
            style="Panel.TRadiobutton",
            command=self._on_mode_changed,
        )
        rb_channel.pack(side=tk.LEFT)
        Tooltip(rb_channel,
                "URLのチャンネルから複数動画のコメントを取得。\n"
                "「リサーチする動画数」で件数を指定できます。")

        # 区切り
        tk.Frame(inner, bg=C_BORDER_SOFT, height=1).pack(
            fill=tk.X, pady=(22, 0))

        # === Section 4: 抽出条件 ===
        ttk.Label(inner, text="抽出条件",
                  style="FieldLabel.TLabel").pack(anchor="w", pady=(18, 12))

        r3 = ttk.Frame(inner, style="Panel.TFrame")
        r3.pack(fill=tk.X, pady=(0, 0))
        # ヘッダー行は上のセクションラベルで代替するため空
        _dummy = r3

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

        # 含まれる文字 + AND/OR モード
        f3 = ttk.Frame(r3b, style="Panel.TFrame")
        f3.pack(side=tk.LEFT, padx=(0, 24))
        ttk.Label(f3, text="コメントに含まれる文字 (カンマ/空白区切りで複数可)",
                  style="Panel.TLabel").pack(anchor="w")
        f3row = ttk.Frame(f3, style="Panel.TFrame")
        f3row.pack(anchor="w", pady=(2, 0))
        self.text_filter_entry = PlaceholderEntry(
            f3row, placeholder="例: :,おもしろ,感動 (カンマ区切り)",
            width=28, font=("", 10))
        self.text_filter_entry.pack(side=tk.LEFT)
        self.text_filter_mode_var = tk.StringVar(value="AND")
        self.text_filter_mode_combo = ttk.Combobox(
            f3row, textvariable=self.text_filter_mode_var,
            values=["AND", "OR"], state="readonly", width=5, font=("", 10),
        )
        self.text_filter_mode_combo.pack(side=tk.LEFT, padx=(4, 0))
        Tooltip(self.text_filter_mode_combo,
                "AND: 全てのキーワードを含むコメントのみ (例: ':' と '感動' の両方含む)\n"
                "OR : いずれかのキーワードを含むコメント (例: ':' か '感動' のどちらか)")
        Tooltip(self.text_filter_entry,
                "複数のキーワードはカンマ(,)または空白で区切ります。\n"
                "AND/OR で条件を切り替えられます。\n"
                "空欄でフィルタなし (全コメント対象)。\n"
                "例: ':' と入力 → コロンを含むコメント (タイムスタンプ付き)\n"
                "例: 'おもしろ,感動' + OR → どちらかを含むコメント")

        # 投稿者名フィルタ
        f4 = ttk.Frame(r3b, style="Panel.TFrame")
        f4.pack(side=tk.LEFT, padx=(0, 24))
        ttk.Label(f4, text="投稿者名に含まれる文字",
                  style="Panel.TLabel").pack(anchor="w")
        self.author_filter_entry = PlaceholderEntry(
            f4, placeholder="例: @channel",
            width=16, font=("", 10))
        self.author_filter_entry.pack(anchor="w", pady=(2, 0))
        Tooltip(self.author_filter_entry,
                "投稿者名にこの文字列を含むコメントのみを出力します (大文字小文字を区別しない)。\n"
                "空欄でフィルタなし。")

        # === Row 3c: 日付範囲フィルタ + 返信・重複除去オプション ===
        r3c = ttk.Frame(inner, style="Panel.TFrame")
        r3c.pack(fill=tk.X, pady=(10, 0))

        # 日付 From / To
        f_date = ttk.Frame(r3c, style="Panel.TFrame")
        f_date.pack(side=tk.LEFT, padx=(0, 24))
        ttk.Label(f_date, text="コメント投稿日の範囲 (YYYY-MM-DD)",
                  style="Panel.TLabel").pack(anchor="w")
        f_date_row = ttk.Frame(f_date, style="Panel.TFrame")
        f_date_row.pack(anchor="w", pady=(2, 0))
        self.date_from_entry = PlaceholderEntry(
            f_date_row, placeholder="From",
            width=12, font=("", 10))
        self.date_from_entry.pack(side=tk.LEFT)
        ttk.Label(f_date_row, text="〜", style="Panel.TLabel").pack(
            side=tk.LEFT, padx=4)
        self.date_to_entry = PlaceholderEntry(
            f_date_row, placeholder="To",
            width=12, font=("", 10))
        self.date_to_entry.pack(side=tk.LEFT)
        Tooltip(self.date_from_entry,
                "コメント投稿日の下限 (この日以降)\n"
                "形式: YYYY-MM-DD (例: 2024-01-01)\n"
                "空欄で下限なし")
        Tooltip(self.date_to_entry,
                "コメント投稿日の上限 (この日まで)\n"
                "形式: YYYY-MM-DD (例: 2024-12-31)\n"
                "空欄で上限なし")

        # チェックボックス群
        f_opts = ttk.Frame(r3c, style="Panel.TFrame")
        f_opts.pack(side=tk.LEFT, padx=(0, 24))
        ttk.Label(f_opts, text="オプション",
                  style="Panel.TLabel").pack(anchor="w")
        f_opts_row = ttk.Frame(f_opts, style="Panel.TFrame")
        f_opts_row.pack(anchor="w", pady=(2, 0))

        self.include_replies_var = tk.BooleanVar(value=False)
        cb_replies = ttk.Checkbutton(
            f_opts_row, text="返信コメントも取得",
            variable=self.include_replies_var,
            style="Panel.TCheckbutton")
        cb_replies.pack(side=tk.LEFT, padx=(0, 12))
        Tooltip(cb_replies,
                "トップレベルのコメントだけでなく、それに対する返信コメントも取得します。\n"
                "※APIクォータの追加消費はありませんが、データ量は増えます。")

        self.dedup_var = tk.BooleanVar(value=True)
        cb_dedup = ttk.Checkbutton(
            f_opts_row, text="重複コメントを除去",
            variable=self.dedup_var,
            style="Panel.TCheckbutton")
        cb_dedup.pack(side=tk.LEFT)
        Tooltip(cb_dedup,
                "同じテキストのコメント (コピペ・スパム等) を除外します。\n"
                "投稿者名+本文の組み合わせで重複判定します。")

        # 区切り
        tk.Frame(inner, bg=C_BORDER_SOFT, height=1).pack(
            fill=tk.X, pady=(22, 0))

        # === Run / Stop ボタン (右下に Apple Primary Button 風) ===
        r4 = ttk.Frame(inner, style="Panel.TFrame")
        r4.pack(fill=tk.X, pady=(20, 0))

        ttk.Label(r4,
                  text="F5 稼働 · Esc 停止 · Ctrl+S 保存 · Ctrl+H 履歴",
                  style="PanelSub.TLabel"
                  ).pack(side=tk.LEFT)

        self.stop_button = ttk.Button(r4, text="停止", style="Stop.TButton",
                                       command=self.stop_extraction,
                                       state=tk.DISABLED)
        self.stop_button.pack(side=tk.RIGHT, padx=(8, 0))

        self.run_button = ttk.Button(r4, text="稼働", style="Primary.TButton",
                                      command=self.start_extraction)
        self.run_button.pack(side=tk.RIGHT)
        Tooltip(self.run_button, "コメント抽出を開始 (F5)")

    # ------------------------------------------------------------------
    # Results section
    # ------------------------------------------------------------------

    def _build_results_section(self):
        outer = ttk.Frame(self.root, padding=(28, 16, 28, 14))
        outer.pack(fill=tk.BOTH, expand=True)

        # ===== ダッシュボード (統計カード) =====
        self.dashboard_frame = ttk.Frame(outer, style="TFrame")
        self.dashboard_frame.pack(fill=tk.X, pady=(0, 6))
        self._dash_cards = {}
        # カード定義: (key, ラベル, emoji)
        card_defs = [
            ("videos",      "動画",       ""),
            ("comments",    "コメント",   ""),
            ("avg_likes",   "平均 ♥",    ""),
            ("max_likes",   "最高 ♥",    ""),
            ("top_author",  "最多投稿者", ""),
            ("replies",     "返信",       ""),
        ]
        _val_font = ("Yu Gothic UI Semibold", 20) if sys.platform == "win32" else ("", 20, "bold")
        for i, (key, label, icon) in enumerate(card_defs):
            # Notion風 極薄境界のカード
            card = tk.Frame(self.dashboard_frame, bg=C_PANEL,
                            highlightbackground=C_BORDER,
                            highlightthickness=1, bd=0)
            card.pack(side=tk.LEFT, fill=tk.X, expand=True,
                      padx=(0 if i == 0 else 10, 0), ipadx=18, ipady=14)
            # ラベル (小さくミュート色)
            ttk.Label(card, text=label,
                      background=C_PANEL, foreground=C_MUTED,
                      font=("Yu Gothic UI", 9) if sys.platform == "win32" else ("", 9)
                      ).pack(anchor="w")
            value_var = tk.StringVar(value="—")
            vl = ttk.Label(card, textvariable=value_var,
                           background=C_PANEL, foreground=C_TEXT,
                           font=_val_font)
            vl.pack(anchor="w", pady=(4, 0))
            self._dash_cards[key] = value_var

        # ヘッダー (件数 + 保存ボタン)
        bar = ttk.Frame(outer)
        bar.pack(fill=tk.X, pady=(10, 6))

        lbl = ttk.Label(bar, text="結果",
                        background=C_BG, foreground=C_TEXT,
                        font=("Yu Gothic UI Semibold", 13) if sys.platform == "win32" else ("", 13, "bold"))
        lbl.pack(side=tk.LEFT)

        # 件数表示 (Notion風 muted)
        self.count_label = ttk.Label(bar, text="0 件",
                                     background=C_BG, foreground=C_MUTED,
                                     font=("Yu Gothic UI", 10) if sys.platform == "win32" else ("", 10))
        self.count_label.pack(side=tk.LEFT, padx=(12, 0))

        # 前回稼働の実消費クォータ表示
        self.quota_result_var = tk.StringVar(value="")
        self.quota_result_label = ttk.Label(bar, textvariable=self.quota_result_var,
                                            background=C_BG, foreground=C_MUTED,
                                            font=("Yu Gothic UI", 9) if sys.platform == "win32" else ("", 9))
        self.quota_result_label.pack(side=tk.LEFT, padx=(16, 0))
        Tooltip(self.quota_result_label,
                "直前の稼働で実際に消費したYouTube APIクォータ量 (実測値)。\n"
                "実残量は Google Cloud Console でご確認ください。")

        # 結果内検索ボックス
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(bar, textvariable=self.search_var, width=24,
                                      font=("", 10))
        self.search_entry.pack(side=tk.LEFT, padx=(24, 0))
        self.search_entry.bind("<KeyRelease>", self._on_search_changed)
        Tooltip(self.search_entry,
                "取得済みの結果から絞り込み検索 (タイトル/投稿者/本文)。")
        # プレースホルダー風挙動 (フォーカス時に淡いラベル)
        ttk.Label(bar, text="", background=C_BG, foreground=C_MUTED).pack(side=tk.LEFT)

        # 保存ボタン (メイン)
        self.export_btn = ttk.Button(bar, text="保存",
                                      style="Primary.TButton",
                                      command=self.save_results, state=tk.DISABLED)
        self.export_btn.pack(side=tk.RIGHT)
        Tooltip(self.export_btn,
                "現在の結果を任意の場所に保存 (Ctrl+S)\n"
                "Excel / CSV / TSV / JSON 対応")

        self.export_csv_btn = ttk.Button(bar, text="CSV",
                                          style="Secondary.TButton",
                                          command=self.save_to_csv, state=tk.DISABLED)
        self.export_csv_btn.pack(side=tk.RIGHT, padx=(0, 8))
        Tooltip(self.export_csv_btn,
                "結果をCSVファイルとして保存 (UTF-8 BOM / Excel対応)")

        clear_btn = ttk.Button(bar, text="クリア",
                               style="Toolbar.TButton",
                               command=self._clear_results)
        clear_btn.pack(side=tk.RIGHT, padx=(0, 8))
        Tooltip(clear_btn, "表示されている結果をクリアします")

        # ===== コンテンツエリア (左: Treeview / 右: プレビュー) =====
        content = ttk.Frame(outer)
        content.pack(fill=tk.BOTH, expand=True)

        # Treeview
        panel = tk.Frame(content, bg="white", highlightbackground=C_BORDER,
                         highlightthickness=1, bd=0)
        panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ===== 右側プレビューパネル (動画サムネ + メタ情報) =====
        self._build_preview_panel(content)

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
        # 列クリックでソート可能にする
        self._sort_state = {}  # {col_id: reverse_bool}
        for cid, heading, width in zip(COL_IDS, HEADERS, COL_WIDTHS_GUI):
            self.tree.heading(
                cid, text=heading,
                command=lambda c=cid: self._sort_by_column(c),
            )
            self.tree.column(cid, width=width, minwidth=50,
                             anchor=anchors.get(cid, "w"))

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_sb.grid(row=0, column=1, sticky="ns")
        x_sb.grid(row=1, column=0, sticky="ew")
        panel.grid_rowconfigure(0, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        # ゼブラストライプ
        # Apple/Notion風 極薄ストライプ
        self.tree.tag_configure("even", background="#fafafa")
        self.tree.tag_configure("odd", background="#ffffff")

        # イベント: ダブルクリックで詳細表示
        self.tree.bind("<Double-1>", self._show_row_detail)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_selection)

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
        self.count_label.config(text="0 件")
        self.export_btn.config(state=tk.DISABLED)
        if hasattr(self, "export_csv_btn"):
            self.export_csv_btn.config(state=tk.DISABLED)
        self.quota_result_var.set("")
        self._reply_count = 0
        if hasattr(self, "_update_dashboard"):
            self._update_dashboard()
        self.progress_var.set(0)
        self.status_var.set("準備完了")

    # ------------------------------------------------------------------
    # Status bar
    # ------------------------------------------------------------------

    def _build_status_bar(self):
        # 上部の細い区切り
        tk.Frame(self.root, bg=C_BORDER_SOFT, height=1).pack(fill=tk.X)

        frame = ttk.Frame(self.root, padding=(28, 10, 28, 14))
        frame.pack(fill=tk.X)

        # プログレスバー (細い / ミニマル)
        pframe = ttk.Frame(frame)
        pframe.pack(fill=tk.X)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            pframe, variable=self.progress_var, maximum=100,
            style="Horizontal.TProgressbar",
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.pct_label = ttk.Label(pframe, text="",
                                   background=C_BG, foreground=C_MUTED,
                                   font=("Yu Gothic UI", 9) if sys.platform == "win32" else ("", 9),
                                   width=6)
        self.pct_label.pack(side=tk.LEFT, padx=(10, 0))

        # ステータステキスト
        sframe = ttk.Frame(frame)
        sframe.pack(fill=tk.X, pady=(8, 0))

        self.status_var = tk.StringVar(value="準備完了")
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
            if hasattr(self, "export_csv_btn"):
                self.export_csv_btn.config(state=tk.DISABLED)
        else:
            self.run_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            if self.results:
                self.export_btn.config(state=tk.NORMAL)
                if hasattr(self, "export_csv_btn"):
                    self.export_csv_btn.config(state=tk.NORMAL)

    def _update_dashboard(self):
        """ダッシュボードカードの数値を再計算して表示。"""
        if not hasattr(self, "_dash_cards"):
            return
        results = self.results
        if not results:
            for k in self._dash_cards:
                self._dash_cards[k].set("—")
            return

        # 動画数 (重複除外)
        video_urls = set(r[0] for r in results)
        # コメント数
        n_comments = len(results)
        # 高評価統計
        likes_list = []
        for r in results:
            try:
                likes_list.append(int(r[5]))
            except (ValueError, TypeError):
                pass
        avg_likes = sum(likes_list) / len(likes_list) if likes_list else 0
        max_likes = max(likes_list) if likes_list else 0
        # 投稿者統計
        author_counts = {}
        for r in results:
            a = r[3]
            if a:
                author_counts[a] = author_counts.get(a, 0) + 1
        top_author = ""
        top_count = 0
        if author_counts:
            top_author, top_count = max(author_counts.items(), key=lambda kv: kv[1])

        # 返信数 (7列目以降に is_reply が入っていれば)
        # 現状の row は (url, title, date, author, text, likes, comment_date) なので0
        replies_count = getattr(self, "_reply_count", 0)

        self._dash_cards["videos"].set(f"{len(video_urls):,}")
        self._dash_cards["comments"].set(f"{n_comments:,}")
        self._dash_cards["avg_likes"].set(f"{avg_likes:,.1f}")
        self._dash_cards["max_likes"].set(f"{max_likes:,}")
        self._dash_cards["top_author"].set(
            f"{(top_author[:14] + '…') if len(top_author) > 14 else top_author}"
            + (f" ({top_count})" if top_count else "")
        )
        self._dash_cards["replies"].set(f"{replies_count:,}")

    def _on_search_changed(self, _evt=None):
        """結果テーブル内のインクリメンタル絞り込み検索。"""
        query = self.search_var.get().strip().lower()
        # 全件をいったん削除して再描画
        self.tree.delete(*self.tree.get_children())
        if not query:
            # 全件表示
            for row in self.results:
                self._add_row_to_tree(row)
            self.count_label.config(text=f"{len(self.results):,} 件")
            return

        visible = 0
        for row in self.results:
            title = str(row[1]).lower()
            author = str(row[3]).lower()
            text = str(row[4]).lower()
            if query in title or query in author or query in text:
                self._add_row_to_tree(row)
                visible += 1

        self.count_label.config(
            text=f"{visible:,} / {len(self.results):,} 件"
        )

    # ------------------------------------------------------------------
    # 履歴機能 (過去の抽出ログ)
    # ------------------------------------------------------------------

    HISTORY_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else ".",
        "history.json",
    )
    MAX_HISTORY = 50

    def _load_history(self) -> list:
        """履歴ファイルを読み込み、リストとして返す。"""
        path = self.HISTORY_PATH
        if not os.path.exists(path):
            return []
        try:
            import json
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        except Exception:
            return []

    def _save_history_entry(self, entry: dict):
        """履歴にエントリを追加して保存 (最新順、最大MAX_HISTORY件)。"""
        try:
            import json
            history = self._load_history()
            history.insert(0, entry)
            history = history[:self.MAX_HISTORY]
            with open(self.HISTORY_PATH, "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
        except Exception as exc:
            # 履歴保存失敗は処理を止めない
            print(f"[warn] 履歴の保存に失敗: {exc}")

    def show_history(self):
        """履歴ダイアログを表示。"""
        history = self._load_history()
        dlg = tk.Toplevel(self.root)
        dlg.title("抽出履歴")
        dlg.geometry("900x500")
        dlg.transient(self.root)

        ttk.Label(dlg, text="📜 過去の抽出履歴 (最大50件)",
                  font=("Yu Gothic UI Semibold", 12) if sys.platform == "win32" else ("", 12, "bold"),
                  ).pack(anchor="w", padx=12, pady=(10, 6))

        if not history:
            ttk.Label(dlg, text="履歴がまだありません。\n稼働すると自動的に記録されます。",
                      foreground=C_MUTED).pack(pady=30)
            ttk.Button(dlg, text="閉じる", command=dlg.destroy).pack(pady=6)
            return

        # Treeviewで履歴リスト
        cols = ("date", "mode", "urls", "comments", "quota")
        tv = ttk.Treeview(dlg, columns=cols, show="headings",
                          style="Results.Treeview", height=14)
        tv.heading("date", text="日時")
        tv.heading("mode", text="モード")
        tv.heading("urls", text="URL / 設定")
        tv.heading("comments", text="コメント数")
        tv.heading("quota", text="消費API")
        tv.column("date", width=140, anchor="w")
        tv.column("mode", width=80, anchor="center")
        tv.column("urls", width=420, anchor="w")
        tv.column("comments", width=80, anchor="center")
        tv.column("quota", width=80, anchor="center")

        for i, h in enumerate(history):
            urls = h.get("urls", [])
            url_str = urls[0] if urls else ""
            if len(urls) > 1:
                url_str += f"  (他{len(urls)-1}件)"
            filters = []
            if h.get("min_likes", 0):
                filters.append(f"♥≥{h['min_likes']}")
            if h.get("text_filter"):
                filters.append(f"'{h['text_filter']}'")
            if h.get("date_from") or h.get("date_to"):
                filters.append("日付範囲")
            summary = url_str
            if filters:
                summary += "  [" + ", ".join(filters) + "]"

            mode_label = "動画のみ" if h.get("mode") == "video" else "チャンネル"
            tv.insert("", tk.END, iid=str(i), values=(
                h.get("timestamp", ""),
                mode_label,
                summary,
                f"{h.get('comment_count', 0):,}",
                f"{h.get('quota', 0):,}",
            ))

        sb = ttk.Scrollbar(dlg, orient=tk.VERTICAL, command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0), pady=(0, 10))
        sb.pack(side=tk.LEFT, fill=tk.Y, pady=(0, 10))

        # ボタンエリア
        btnf = ttk.Frame(dlg)
        btnf.pack(side=tk.RIGHT, fill=tk.Y, padx=12, pady=(0, 10))

        def _reload_selected():
            sel = tv.selection()
            if not sel:
                return
            idx = int(sel[0])
            h = history[idx]
            # 条件を入力欄に復元
            self._clear_all_urls()
            for url in h.get("urls", []):
                if self.url_entries and not self.url_entries[0].get_value().strip():
                    self.url_entries[0].set_value(url)
                else:
                    self._add_url_row(url)
            self.extract_mode_var.set(h.get("mode", "video"))
            self._on_mode_changed()
            self.max_videos_var.set(str(h.get("max_videos", 30)))
            self.min_likes_var.set(str(h.get("min_likes", 0)))
            self.text_filter_entry.set_value(h.get("text_filter", ""))
            self.text_filter_mode_var.set(h.get("text_filter_mode", "AND"))
            self.author_filter_entry.set_value(h.get("author_filter", ""))
            self.include_replies_var.set(h.get("include_replies", False))
            self.dedup_var.set(h.get("dedup", True))
            # 並び順
            order_val = h.get("order", "date")
            self.order_var.set(order_val)
            label = self._order_value_to_label.get(order_val, "新しい順")
            self.order_display_var.set(label)
            # 日付
            if h.get("date_from"):
                self.date_from_entry.set_value(h["date_from"][:10])
            if h.get("date_to"):
                self.date_to_entry.set_value(h["date_to"][:10])
            dlg.destroy()
            messagebox.showinfo(
                "条件を復元しました",
                "履歴から抽出条件を復元しました。\n"
                "APIキーとご確認の上、「▶ 稼働」をクリックしてください。"
            )

        def _open_file():
            sel = tv.selection()
            if not sel:
                return
            h = history[int(sel[0])]
            fp = h.get("export_path", "")
            if fp and os.path.exists(fp):
                try:
                    if sys.platform == "win32":
                        os.startfile(fp)
                    elif sys.platform == "darwin":
                        os.system(f'open "{fp}"')
                    else:
                        os.system(f'xdg-open "{fp}"')
                except Exception as e:
                    messagebox.showerror("エラー", f"ファイルを開けません: {e}")
            else:
                messagebox.showinfo(
                    "情報",
                    "保存ファイルが見つかりません。\n"
                    "削除されたか、別の場所に移動された可能性があります。"
                )

        def _delete_selected():
            sel = tv.selection()
            if not sel:
                return
            if not messagebox.askyesno("確認", "選択した履歴を削除しますか?"):
                return
            idx = int(sel[0])
            history.pop(idx)
            try:
                import json
                with open(self.HISTORY_PATH, "w", encoding="utf-8") as f:
                    json.dump(history, f, ensure_ascii=False, indent=2)
            except Exception as e:
                messagebox.showerror("エラー", f"削除に失敗しました: {e}")
            tv.delete(sel[0])

        def _clear_all():
            if not messagebox.askyesno("確認", "全ての履歴を削除しますか? この操作は元に戻せません。"):
                return
            try:
                if os.path.exists(self.HISTORY_PATH):
                    os.remove(self.HISTORY_PATH)
                dlg.destroy()
                messagebox.showinfo("完了", "履歴を全て削除しました。")
            except Exception as e:
                messagebox.showerror("エラー", f"削除に失敗しました: {e}")

        ttk.Button(btnf, text="🔄 この条件で再実行",
                   style="Primary.TButton",
                   command=_reload_selected).pack(fill=tk.X, pady=(0, 6))
        ttk.Button(btnf, text="📂 保存ファイルを開く",
                   command=_open_file).pack(fill=tk.X, pady=3)
        ttk.Button(btnf, text="🗑 この履歴を削除",
                   command=_delete_selected).pack(fill=tk.X, pady=3)
        ttk.Button(btnf, text="全て削除",
                   command=_clear_all).pack(fill=tk.X, pady=(12, 3))
        ttk.Button(btnf, text="閉じる",
                   command=dlg.destroy).pack(fill=tk.X, pady=(12, 0))

    # ------------------------------------------------------------------
    # プレビューパネル (動画サムネイル + メタ情報)
    # ------------------------------------------------------------------

    def _build_preview_panel(self, parent):
        """右側に動画プレビューパネル。Apple風のミニマルデザイン。"""
        self._thumb_cache = {}  # video_id -> PhotoImage
        self._thumb_lock = threading.Lock()

        # Notion風 極薄境界のパネル
        preview = tk.Frame(parent, bg=C_PANEL,
                           highlightbackground=C_BORDER,
                           highlightthickness=1, bd=0,
                           width=300)
        preview.pack(side=tk.LEFT, fill=tk.Y, padx=(16, 0))
        preview.pack_propagate(False)

        # 見出し (ミュート色)
        ttk.Label(preview, text="PREVIEW",
                  background=C_PANEL, foreground=C_MUTED,
                  font=("Yu Gothic UI", 9) if sys.platform == "win32" else ("", 9),
                  ).pack(anchor="w", padx=20, pady=(18, 4))

        # サムネイルエリア (角丸っぽい印象)
        self.thumb_label = tk.Label(preview, bg=C_BG_SUBTLE,
                                     width=260, height=146, text="行を選択",
                                     fg=C_MUTED, bd=0,
                                     font=("Yu Gothic UI", 9) if sys.platform == "win32" else ("", 9))
        self.thumb_label.pack(padx=20, pady=(0, 14))

        # 動画タイトル
        self.preview_title_var = tk.StringVar(value="")
        ttk.Label(preview, textvariable=self.preview_title_var,
                  background=C_PANEL, foreground=C_TEXT,
                  wraplength=260,
                  font=("Yu Gothic UI Semibold", 11) if sys.platform == "win32" else ("", 11, "bold"),
                  ).pack(anchor="w", padx=20, pady=(0, 6))

        self.preview_meta_var = tk.StringVar(value="")
        ttk.Label(preview, textvariable=self.preview_meta_var,
                  background=C_PANEL, foreground=C_MUTED, wraplength=260,
                  font=("Yu Gothic UI", 9) if sys.platform == "win32" else ("", 9),
                  ).pack(anchor="w", padx=20, pady=(0, 14))

        # YouTubeで開くボタン (Link style)
        self._preview_url = None
        btn = ttk.Button(preview, text="YouTubeで開く →",
                         style="Link.TButton",
                         command=self._open_preview_video)
        btn.pack(padx=20, pady=(0, 16), anchor="w")

        if not PIL_AVAILABLE:
            tk.Frame(preview, bg=C_BORDER_SOFT, height=1).pack(
                fill=tk.X, padx=20, pady=(0, 10))
            ttk.Label(preview,
                      text="サムネイル表示には Pillow が必要です",
                      background=C_PANEL, foreground=C_MUTED,
                      wraplength=260,
                      font=("Yu Gothic UI", 8) if sys.platform == "win32" else ("", 8),
                      ).pack(padx=20, pady=(0, 12))

    def _open_preview_video(self):
        if self._preview_url:
            webbrowser.open(self._preview_url)

    def _on_tree_selection(self, _evt=None):
        """Treeview選択変更 → プレビューパネル更新。"""
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if not vals or len(vals) < 7:
            return

        video_url, title, vid_date, author, _text, _likes, _cdate = vals
        self._preview_url = str(video_url)

        self.preview_title_var.set(title)
        # コメント数をカウント
        n_comments = sum(1 for r in self.results if r[0] == video_url)
        self.preview_meta_var.set(
            f"📅 投稿日: {vid_date}\n"
            f"💬 取得コメント数: {n_comments} 件"
        )

        # サムネイルを非同期ロード
        video_id = self._extract_video_id(str(video_url))
        if video_id and PIL_AVAILABLE:
            self._load_thumbnail_async(video_id)

    def _extract_video_id(self, url: str) -> str:
        import re
        m = re.search(r"[?&]v=([a-zA-Z0-9_-]+)", url)
        if m:
            return m.group(1)
        m = re.search(r"/shorts/([a-zA-Z0-9_-]+)", url)
        if m:
            return m.group(1)
        m = re.search(r"/live/([a-zA-Z0-9_-]+)", url)
        if m:
            return m.group(1)
        return ""

    def _load_thumbnail_async(self, video_id: str):
        """YouTubeのサムネイルをダウンロードして表示 (キャッシュあり)。"""
        if video_id in self._thumb_cache:
            self._display_thumbnail(self._thumb_cache[video_id])
            return

        self.thumb_label.configure(image="", text="読み込み中...", fg="#aaa")

        def _worker():
            # 複数解像度を試す (mqdefault = 320x180, hqdefault = 480x360)
            urls = [
                f"https://i.ytimg.com/vi/{video_id}/mqdefault.jpg",
                f"https://i.ytimg.com/vi/{video_id}/hqdefault.jpg",
                f"https://i.ytimg.com/vi/{video_id}/default.jpg",
            ]
            for url in urls:
                try:
                    with urllib.request.urlopen(url, timeout=5) as resp:
                        data = resp.read()
                    from io import BytesIO
                    img = Image.open(BytesIO(data))
                    img.thumbnail((256, 144))
                    photo = ImageTk.PhotoImage(img)
                    with self._thumb_lock:
                        self._thumb_cache[video_id] = photo
                    # GUI更新はメインスレッドで
                    self.root.after(0, lambda: self._display_thumbnail(photo, check_id=video_id))
                    return
                except Exception:
                    continue
            # すべて失敗
            self.root.after(0, lambda: self.thumb_label.configure(
                image="", text="(サムネイル取得失敗)", fg="#aaa"))

        threading.Thread(target=_worker, daemon=True).start()

    def _display_thumbnail(self, photo, check_id: str = None):
        """サムネイル画像をプレビューに表示 (選択が変わっていたら無視)。"""
        # 表示時に別の動画が選択されていたら無視
        if check_id:
            current_id = self._extract_video_id(self._preview_url or "")
            if current_id != check_id:
                return
        self.thumb_label.configure(image=photo, text="")
        self.thumb_label.image = photo  # GC防止

    def _sort_by_column(self, col_id: str):
        """列ヘッダークリックで結果テーブルをソート (昇順⇔降順トグル)。"""
        if not self.tree.get_children():
            return
        # 型別ソートキー
        def sort_key(row_id):
            val = self.tree.set(row_id, col_id)
            if col_id == "likes":
                try:
                    return (0, int(val))
                except ValueError:
                    return (0, 0)
            # 日付/日時は文字列比較で元のフォーマットなら正常動作
            return (1, val.lower())

        reverse = self._sort_state.get(col_id, False)
        rows = [(sort_key(r), r) for r in self.tree.get_children()]
        rows.sort(key=lambda x: x[0], reverse=reverse)
        for i, (_, r) in enumerate(rows):
            self.tree.move(r, "", i)
            # ゼブラ再適用
            self.tree.item(r, tags=("even" if i % 2 == 0 else "odd",))

        # トグル
        self._sort_state[col_id] = not reverse

        # 見出しに矢印アイコン
        for cid, heading in zip(COL_IDS, HEADERS):
            if cid == col_id:
                arrow = " ▼" if reverse else " ▲"
                self.tree.heading(cid, text=heading + arrow)
            else:
                self.tree.heading(cid, text=heading)

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
                    # 絞り込み中の場合は検索条件に合致する場合のみ追加表示
                    query = self.search_var.get().strip().lower() if hasattr(self, "search_var") else ""
                    if not query or (query in str(msg["data"][1]).lower() or
                                     query in str(msg["data"][3]).lower() or
                                     query in str(msg["data"][4]).lower()):
                        self._add_row_to_tree(msg["data"])
                    count_txt = f"{len(self.results):,} 件"
                    if query:
                        visible = len(self.tree.get_children())
                        count_txt = f"{visible:,} / {len(self.results):,} 件"
                    self.count_label.config(text=count_txt)
                    # ダッシュボードは一定間隔で更新 (頻繁すぎる更新を抑制)
                    self._dashboard_dirty = True
                elif kind == "done":
                    self._on_done(msg.get("error"))
        except queue.Empty:
            pass
        # ダッシュボードが汚れていたら更新 (メッセージ処理後1回)
        if getattr(self, "_dashboard_dirty", False):
            self._update_dashboard()
            self._dashboard_dirty = False

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
        text_filter_mode = self.text_filter_mode_var.get()
        author_filter = self.author_filter_entry.get_value().strip()
        include_replies = self.include_replies_var.get()
        dedup = self.dedup_var.get()

        # 日付範囲: YYYY-MM-DD 形式を ISO 8601 に変換
        date_from_str = self.date_from_entry.get_value().strip()
        date_to_str = self.date_to_entry.get_value().strip()
        date_from = None
        date_to = None
        try:
            if date_from_str:
                # From はその日の0時として扱う
                datetime.strptime(date_from_str, "%Y-%m-%d")
                date_from = f"{date_from_str}T00:00:00Z"
            if date_to_str:
                datetime.strptime(date_to_str, "%Y-%m-%d")
                # To はその日の23:59:59として扱う
                date_to = f"{date_to_str}T23:59:59Z"
        except ValueError:
            messagebox.showwarning(
                "入力エラー",
                "日付は YYYY-MM-DD 形式で入力してください (例: 2024-01-15)"
            )
            return

        # 結果をクリア
        self.tree.delete(*self.tree.get_children())
        self.results.clear()
        self._seen_comment_keys = set()  # 重複除去用
        self._reply_count = 0
        self.count_label.config(text="0 件")
        if hasattr(self, "_update_dashboard"):
            self._update_dashboard()
        self._update_progress(0)
        self.elapsed_var.set("")

        self.cancelled = False
        self._start_time = time.time()
        self._set_running_state(True)
        self._update_elapsed()

        # 現在実行中の設定を保存 (履歴用)
        self._current_params = {
            "urls": urls,
            "mode": mode,
            "order": order,
            "max_videos": max_videos,
            "min_likes": min_likes,
            "text_filter": text_filter,
            "text_filter_mode": text_filter_mode,
            "author_filter": author_filter,
            "include_replies": include_replies,
            "dedup": dedup,
            "date_from": date_from,
            "date_to": date_to,
        }

        self.worker_thread = threading.Thread(
            target=self._worker,
            kwargs={
                "api_key": api_key,
                "urls": urls,
                "max_videos": max_videos,
                "min_likes": min_likes,
                "text_filter": text_filter,
                "text_filter_mode": text_filter_mode,
                "author_filter": author_filter,
                "include_replies": include_replies,
                "dedup": dedup,
                "date_from": date_from,
                "date_to": date_to,
                "mode": mode,
                "order": order,
            },
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
                text_filter_mode="AND", author_filter="",
                include_replies=False, dedup=True,
                date_from=None, date_to=None,
                mode="video", order="date"):
        put = self.msg_queue.put

        # --- APIキーローテーション準備 ---
        # メインのapi_keyが先頭、次にapi_keys.jsonの残りキーを続ける
        saved_keys = [k.get("key") for k in self._load_api_keys() if k.get("key")]
        key_pool = [api_key] + [k for k in saved_keys if k and k != api_key]
        key_labels = {api_key: "(入力欄のキー)"}
        for saved in self._load_api_keys():
            if saved.get("key"):
                key_labels.setdefault(saved["key"],
                                       saved.get("label", "(ラベルなし)"))

        # 現在使用中のキーのインデックス
        self._active_key_idx = 0
        client = self._init_client_with_rotation(key_pool, key_labels, put)
        if client is None:
            put({"type": "done", "error": "利用可能なAPIキーがありません。"})
            return

        try:

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
                        text_filter_mode=text_filter_mode,
                        author_filter=author_filter,
                        include_replies=include_replies,
                        date_from=date_from,
                        date_to=date_to,
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

                # 重複除去用: (author, text) で判定
                seen_keys = getattr(self, "_seen_comment_keys", None)
                if seen_keys is None:
                    self._seen_comment_keys = seen_keys = set()

                for c in comments:
                    if self.cancelled:
                        break

                    # 重複除去 (author + text)
                    if dedup:
                        key = (c["author"], c["text"])
                        if key in seen_keys:
                            continue
                        seen_keys.add(key)

                    if c.get("is_reply"):
                        self._reply_count = getattr(self, "_reply_count", 0) + 1

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
            self._consumed_quota = client.quota_used
            put({"type": "done", "error": None})

        except APIError as exc:
            # クォータ枯渇 (403) の場合はキー切替を試みる
            if exc.status == 403 and len(key_pool) > 1:
                self._consumed_quota = getattr(client, "quota_used", 0)
                new_client = self._try_rotate_key(client, key_pool, key_labels, put)
                if new_client is not None:
                    # キー切替成功 → 再度実行を試みる
                    put({
                        "type": "status",
                        "text": f"キー切替成功。処理を再開します..."
                    })
                    # 簡易再実行: 再帰的にワーカーを呼び出す
                    return self._worker(
                        api_key=key_pool[self._active_key_idx],
                        urls=urls, max_videos=max_videos,
                        min_likes=min_likes, text_filter=text_filter,
                        text_filter_mode=text_filter_mode,
                        author_filter=author_filter,
                        include_replies=include_replies, dedup=dedup,
                        date_from=date_from, date_to=date_to,
                        mode=mode, order=order,
                    )
            try:
                self._consumed_quota = client.quota_used
            except (NameError, AttributeError):
                self._consumed_quota = 0
            put({"type": "done", "error": f"APIエラー: {exc}"})
        except ValueError as exc:
            try:
                self._consumed_quota = client.quota_used
            except (NameError, AttributeError):
                self._consumed_quota = 0
            put({"type": "done", "error": str(exc)})
        except Exception as exc:
            try:
                self._consumed_quota = client.quota_used
            except (NameError, AttributeError):
                self._consumed_quota = 0
            put({"type": "done", "error": f"予期しないエラー: {type(exc).__name__}: {exc}"})

    # ------------------------------------------------------------------
    # Completion
    # ------------------------------------------------------------------

    def _on_done(self, error):
        self._set_running_state(False)

        # 実消費クォータ (ワーカーから通知された実測値)
        consumed = getattr(self, "_consumed_quota", 0)
        quota_text = f" | 消費API: {consumed:,} units" if consumed else ""

        # 結果エリアの消費クォータ表示を更新
        if hasattr(self, "quota_result_var"):
            if consumed:
                self.quota_result_var.set(f"📊 消費API: {consumed:,} units")
            else:
                self.quota_result_var.set("")

        if error:
            self.status_var.set(f"エラー: {error}{quota_text}")
            messagebox.showerror(
                "エラー",
                f"処理中にエラーが発生しました:\n\n{error}\n\n"
                f"消費したAPIクォータ: {consumed:,} units\n"
                "取得済みのデータはテーブルに保持されています。\n"
                "「💾 Excelに名前を付けて保存」から保存できます。",
            )
        elif self.cancelled:
            self.status_var.set(
                f"⏸ 停止しました (取得済み: {len(self.results)} 件){quota_text}"
            )
        else:
            self.status_var.set(
                f"✓ 完了！ 合計 {len(self.results)} 件のコメントを取得しました{quota_text}"
            )

        # ダッシュボード最終更新
        if hasattr(self, "_update_dashboard"):
            self._update_dashboard()

        if self.results:
            self._auto_save()

    def _auto_save(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = f"youtube_comments_{ts}.xlsx"
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
        saved = False
        try:
            self._write_excel(path)
            cur = self.status_var.get()
            self.status_var.set(f"{cur}  |  自動保存: {name}")
            saved = True
        except Exception as exc:
            messagebox.showwarning(
                "自動保存エラー",
                f"Excelファイルの自動保存に失敗しました:\n{exc}\n\n"
                "「💾 保存...」から手動で保存してください。"
            )

        # 履歴に記録
        try:
            params = getattr(self, "_current_params", {}).copy()
            entry = {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                **params,
                "comment_count": len(self.results),
                "quota": getattr(self, "_consumed_quota", 0),
                "export_path": path if saved else "",
            }
            self._save_history_entry(entry)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Excel export
    # ------------------------------------------------------------------

    def save_results(self):
        """
        保存先を選択し、拡張子に応じて Excel / CSV / TSV / JSON を出力。
        """
        if not self.results:
            messagebox.showinfo(
                "情報",
                "保存するデータがありません。\n先に「稼働」でデータを取得してください。"
            )
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = f"youtube_comments_{ts}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel ファイル", "*.xlsx"),
                ("CSV (UTF-8 BOM付き, Excel互換)", "*.csv"),
                ("TSV (タブ区切り)", "*.tsv"),
                ("JSON", "*.json"),
                ("全てのファイル", "*.*"),
            ],
            initialfile=default,
            title="結果の保存先と形式を選択",
        )
        if not path:
            return

        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in (".csv",):
                self._write_csv(path)
            elif ext in (".tsv",):
                self._write_csv(path, delimiter="\t")
            elif ext in (".json",):
                self._write_json(path)
            else:
                # 既定は Excel
                self._write_excel(path)
            messagebox.showinfo("保存完了", f"保存しました:\n{path}")
        except Exception as exc:
            messagebox.showerror("保存エラー", f"保存に失敗しました:\n{exc}")

    # 後方互換エイリアス
    save_to_excel = save_results

    def save_to_csv(self):
        """CSVファイルとして素早く保存。"""
        if not self.results:
            messagebox.showinfo(
                "情報",
                "保存するデータがありません。\n先に「稼働」でデータを取得してください。"
            )
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = f"youtube_comments_{ts}.csv"
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV (UTF-8 BOM付き, Excel互換)", "*.csv")],
            initialfile=default,
            title="CSVファイルの保存先を選択",
        )
        if not path:
            return
        try:
            self._write_csv(path)
            messagebox.showinfo("保存完了", f"CSVファイルを保存しました:\n{path}")
        except Exception as exc:
            messagebox.showerror("保存エラー", f"保存に失敗しました:\n{exc}")

    def _write_csv(self, filepath: str, delimiter: str = ","):
        """CSV/TSV として結果を書き出す (UTF-8 BOM付き、Excel互換)。"""
        import csv
        # utf-8-sig で BOM付き → Excel で文字化けせず開ける
        with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
            writer.writerow(HEADERS)
            for row in self.results:
                writer.writerow(row)

    def _write_json(self, filepath: str):
        """JSON として結果を書き出す (各行を辞書化)。"""
        import json
        keys = ["video_url", "video_title", "video_published_at",
                "author", "comment_text", "comment_likes", "comment_published_at"]
        payload = {
            "meta": {
                "extracted_at": datetime.now().isoformat(timespec="seconds"),
                "count": len(self.results),
                "filters": {
                    "min_likes": self.min_likes_var.get(),
                    "text_filter": self.text_filter_entry.get_value(),
                },
                "urls": self._get_all_urls(),
                "mode": self.extract_mode_var.get(),
                "order": self.order_var.get() if hasattr(self, "order_var") else "date",
                "consumed_quota": getattr(self, "_consumed_quota", 0),
            },
            "comments": [dict(zip(keys, row)) for row in self.results],
        }
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

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

    def _init_client_with_rotation(self, key_pool, key_labels, put):
        """
        キーのプールから順番に試して、最初に動作するYouTubeClientを返す。
        全キーで失敗した場合は None。
        """
        for idx, k in enumerate(key_pool):
            if not k:
                continue
            try:
                client = YouTubeClient(k)
                # 検証コールを後でまとめて行うため、ここではインスタンス化のみ
                self._active_key_idx = idx
                self._active_key_label = key_labels.get(k, "(ラベルなし)")
                put({
                    "type": "status",
                    "text": f"API接続中... キー: {self._active_key_label}"
                })
                return client
            except Exception as exc:
                put({
                    "type": "status",
                    "text": f"キー #{idx + 1} の初期化失敗: {exc}"
                })
                continue
        return None

    def _try_rotate_key(self, current_client, key_pool, key_labels, put):
        """
        現在のキーが使えなくなった時、次のキーを試す。
        成功したら新しいYouTubeClient、失敗したらNone。
        """
        start_idx = self._active_key_idx + 1
        for idx in range(start_idx, len(key_pool)):
            k = key_pool[idx]
            if not k:
                continue
            try:
                new_client = YouTubeClient(k)
                self._active_key_idx = idx
                self._active_key_label = key_labels.get(k, "(ラベルなし)")
                put({
                    "type": "status",
                    "text": (f"🔁 次のAPIキーに切替: {self._active_key_label}")
                })
                return new_client
            except Exception:
                continue
        return None

    # ------------------------------------------------------------------
    # APIキー管理 (複数キーのローテーション対応)
    # ------------------------------------------------------------------

    API_KEYS_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else ".",
        "api_keys.json",
    )

    def _load_api_keys(self) -> list:
        """api_keys.json から保存済みキーを読み込む。"""
        if not os.path.exists(self.API_KEYS_PATH):
            return []
        try:
            import json
            with open(self.API_KEYS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        except Exception:
            return []

    def _save_api_keys(self, keys: list):
        """api_keys.json にキーリストを保存。"""
        try:
            import json
            with open(self.API_KEYS_PATH, "w", encoding="utf-8") as f:
                json.dump(keys, f, ensure_ascii=False, indent=2)
            return True
        except Exception as exc:
            messagebox.showerror("保存エラー", f"APIキーリストの保存に失敗: {exc}")
            return False

    def show_api_keys_manager(self):
        """APIキー管理ダイアログを表示。"""
        keys = self._load_api_keys()
        dlg = tk.Toplevel(self.root)
        dlg.title("APIキー管理 (ローテーション)")
        dlg.geometry("720x460")
        dlg.transient(self.root)

        ttk.Label(
            dlg,
            text="🔑 APIキー管理",
            font=("Yu Gothic UI Semibold", 13) if sys.platform == "win32" else ("", 13, "bold"),
        ).pack(anchor="w", padx=14, pady=(12, 4))

        ttk.Label(
            dlg,
            text=("複数のYouTube Data APIキーを登録すると、1つがクォータ制限に達した時に自動で次のキーに切り替えます。\n"
                  "※ 複数のGoogleアカウントで取得したキーを利用してください (同じアカウントの複数キーは同じクォータ枠を共有します)。"),
            foreground=C_TEXT_SEC, wraplength=680,
        ).pack(anchor="w", padx=14, pady=(0, 10))

        # キーリスト
        lf = ttk.LabelFrame(dlg, text="登録済みキー", padding=10)
        lf.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 10))

        tv = ttk.Treeview(lf, columns=("label", "key"), show="headings",
                          style="Results.Treeview", height=8)
        tv.heading("label", text="ラベル")
        tv.heading("key", text="APIキー (マスク表示)")
        tv.column("label", width=160)
        tv.column("key", width=440)

        def _refresh():
            tv.delete(*tv.get_children())
            for i, k in enumerate(keys):
                masked = k["key"][:8] + "…" + k["key"][-4:] if len(k["key"]) > 12 else "…"
                tv.insert("", tk.END, iid=str(i),
                          values=(k.get("label", "(ラベルなし)"), masked))

        _refresh()
        tv.pack(fill=tk.BOTH, expand=True)

        # 入力エリア
        input_frame = ttk.Frame(dlg)
        input_frame.pack(fill=tk.X, padx=14, pady=(0, 10))

        ttk.Label(input_frame, text="ラベル:").pack(side=tk.LEFT)
        label_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=label_var, width=14).pack(
            side=tk.LEFT, padx=(4, 10))
        ttk.Label(input_frame, text="APIキー:").pack(side=tk.LEFT)
        key_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=key_var, width=44, show="●").pack(
            side=tk.LEFT, padx=(4, 10), fill=tk.X, expand=True)

        def _add_key():
            k = key_var.get().strip()
            lbl = label_var.get().strip() or f"Key {len(keys) + 1}"
            if not k:
                messagebox.showwarning("入力エラー", "APIキーを入力してください。")
                return
            keys.append({"label": lbl, "key": k})
            if self._save_api_keys(keys):
                _refresh()
                label_var.set("")
                key_var.set("")

        def _delete():
            sel = tv.selection()
            if not sel:
                return
            idx = int(sel[0])
            if messagebox.askyesno("確認", f"キー「{keys[idx].get('label')}」を削除しますか?"):
                keys.pop(idx)
                if self._save_api_keys(keys):
                    _refresh()

        def _use_selected():
            sel = tv.selection()
            if not sel:
                messagebox.showwarning("情報", "使用するキーを選択してください。")
                return
            idx = int(sel[0])
            self.api_key_var.set(keys[idx]["key"])
            dlg.destroy()
            messagebox.showinfo(
                "適用しました",
                f"キー「{keys[idx].get('label')}」をメイン画面にセットしました。"
            )

        ttk.Button(input_frame, text="＋ 追加", style="Primary.TButton",
                   command=_add_key).pack(side=tk.LEFT)

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(fill=tk.X, padx=14, pady=(0, 14))
        ttk.Button(btn_frame, text="✓ 選択したキーを使用",
                   style="Primary.TButton",
                   command=_use_selected).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="🗑 選択したキーを削除",
                   command=_delete).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(btn_frame, text="閉じる",
                   command=dlg.destroy).pack(side=tk.RIGHT)

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
