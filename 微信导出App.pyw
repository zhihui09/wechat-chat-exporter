#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""微信聊天记录导出 - 图形界面

本轮改动：
1. 只列「一对一联系人」：好友 / 陌生人（含已删好友），群聊与公众号不进列表。
2. 按备注名（无备注用昵称）的拼音首字母 A-Z 排序，左侧字母索引可跳转。
3. 好友 / 陌生人 两个分组，可切换。
4. 界面改为圆角卡片式：浅色底、圆角面板、圆角按钮、加大行距。
"""
from __future__ import annotations

import os
import sys
import queue
import subprocess
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

# 最早的日志钩子：在 import 任何重模块之前就写好，这样打包版闪退也能查
# 打包成 单文件/无控制台 后 BASE 可能指向 _MEIPASS 或不可写目录，所以多试几处，
# 并用 O_APPEND 逐行 append（避免多个实例互相截断）。
_EARLY_LOG = None
_early_cands = [
    Path(os.environ.get("TEMP") or r"C:\Windows\Temp") / "WechatExportTool_early.log",
    Path(os.environ.get("LOCALAPPDATA") or "") / "WechatExportTool_early.log",
    BASE / "early.log",
]
for _cand in _early_cands:
    try:
        if not _cand.parent or not _cand.parent.exists():
            continue
        with open(_cand, "a", encoding="utf-8"):
            pass
        _EARLY_LOG = _cand
        break
    except Exception:
        continue


def elog(msg: str) -> None:
    if _EARLY_LOG is None:
        return
    try:
        with open(_EARLY_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%H:%M:%S}] {msg}\n")
    except Exception:
        pass


elog(f"=== 启动 frozen={getattr(sys, 'frozen', False)}")
elog(f"exe={sys.executable}")

APP = "微信聊天记录导出"

# ---------------------------------------------------------------- 日志
def _pick_log_path() -> tuple[Path | None, str]:
    """挑一个真能写的日志位置；全都写不了就返回 None（不写日志）。

    顺序：程序自己旁边 -> %TEMP% -> 放弃。
    任何一步失败都只记录原因，绝不抛异常（上一版就是这里把程序搞死的）。
    """
    cands: list[Path] = []
    try:
        cands.append(BASE / "log.txt")
    except Exception:
        pass
    try:
        tmp = os.environ.get("TEMP") or os.environ.get("TMP")
        if tmp:
            cands.append(Path(tmp) / "WechatExportTool.log")
    except Exception:
        pass
    for c in cands:
        try:
            c.parent.mkdir(parents=True, exist_ok=True)
            with open(c, "a", encoding="utf-8") as f:
                f.write("")
            return c, ""
        except Exception:
            continue
    return None, "所有候选位置都不可写"


LOG_PATH, _LOG_ERR = _pick_log_path()


def log(msg: str) -> None:
    """写日志，永远不抛异常；没有可用路径就静默跳过。"""
    if LOG_PATH is None:
        return
    try:
        line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n"
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def _early_fail(title: str, text: str) -> None:
    """在 tkinter 起来之前也要能把错误说出去。"""
    log(f"EARLY_FAIL {title}: {text}")
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(title, text)
        r.destroy()
    except Exception:
        pass


log("=" * 60)
log(f"启动 pid={os.getpid()} frozen={getattr(sys, 'frozen', False)}")
log(f"executable={sys.executable}")
log(f"argv={sys.argv}")
log(f"cwd={os.getcwd()}")

# ---------------------------------------------------------------- 单实例锁
_MUTEX = None


def acquire_single_instance() -> bool:
    """已有一个实例在跑就返回 False。非 Windows 直接放行。"""
    global _MUTEX
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.restype = wintypes.HANDLE
        kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL,
                                          wintypes.LPCWSTR]
        _MUTEX = kernel32.CreateMutexW(None, True, "Global\\WechatExportTool_Single")
        err = ctypes.get_last_error()
        if err == 183:                      # ERROR_ALREADY_EXISTS
            log("检测到已有实例在运行，本次直接退出")
            return False
        log(f"单实例锁获取成功，last_error={err}")
        return True
    except Exception as e:
        log(f"单实例锁不可用（忽略）：{type(e).__name__}: {e}")
        return True


def focus_existing() -> None:
    """尽量把已有窗口拉到前台。"""
    try:
        import ctypes
        u32 = ctypes.windll.user32
        hwnd = u32.FindWindowW(None, APP)
        if hwnd:
            u32.ShowWindow(hwnd, 9)         # SW_RESTORE
            u32.SetForegroundWindow(hwnd)
            log(f"已聚焦已有窗口 hwnd={hwnd}")
        else:
            log("没找到已有窗口句柄")
    except Exception as e:
        log(f"聚焦失败：{e}")


if not acquire_single_instance():
    focus_existing()
    sys.exit(0)

# ---------------------------------------------------------------- 依赖自检
try:
    import wxexport_core as core
    log("wxexport_core 导入成功")
except Exception as e:
    log(f"wxexport_core 导入失败：{type(e).__name__}: {e}")
    log(traceback.format_exc())
    _early_fail("微信聊天记录导出",
                f"核心模块加载失败：\n{type(e).__name__}: {e}\n\n日志：{LOG_PATH}")
    sys.exit(1)

_missing = core.missing_deps()
log(f"missing_deps = {_missing}")

# 只把「真正被需要的模块」当硬依赖（缺了必须换解释器）。
_HARD = ("docx",)
_hard_missing = [m for m in _missing if m in _HARD]
_soft_missing = [m for m in _missing if m not in _HARD]

if _soft_missing:
    log(f"软缺失（不阻断启动）：{_soft_missing}")

if _hard_missing:
    log(f"硬依赖缺失：{_hard_missing}")
    _py = "" if getattr(sys, "frozen", False) else core.find_python()
    log(f"frozen={getattr(sys, 'frozen', False)} find_python -> {_py!r}")
    if (not getattr(sys, "frozen", False)) and _py and \
            Path(_py).resolve() != Path(sys.executable).resolve():
        log(f"用 {_py} 重启自身")
        os.execv(_py, [_py, str(Path(__file__).resolve()), *sys.argv[1:]])
    _early_fail("微信聊天记录导出",
                "缺少 Python 依赖：" + ", ".join(_hard_missing) +
                f"\n\n日志：{LOG_PATH}")
    sys.exit(1)


# ---------------------------------------------------------------- headless e2e
def _headless_e2e() -> int:
    """打包版端到端验收：不起 GUI，走 prepare→选人→collect→Word/TXT→媒体归集→产物校验。

    参数从 exe 旁 e2e_args.txt 读（第 1 行联系人关键字、可空=选第一个；第 2 行输出目录），
    避免命令行传中文乱码。结果写 log.txt，进程退出码即结论。
    """
    import zipfile
    args = Path(sys.executable).parent / "e2e_args.txt"
    kw, outdir = "", Path.home() / "Documents" / "微信导出" / "e2e"
    try:
        lines = args.read_text(encoding="utf-8").splitlines()
        if lines:
            kw = lines[0].strip()
        if len(lines) > 1 and lines[1].strip():
            outdir = Path(lines[1].strip())
    except OSError:
        pass
    log(f"E2E 开始 kw={kw!r} out={outdir}")
    if not core.is_admin():
        log("E2E FAIL: 非管理员"); return 2
    try:
        outdir.mkdir(parents=True, exist_ok=True)
        ex = core.Exporter(log=lambda m: log("  " + m),
                           progress=lambda c, t, m: None)
        dec = ex.prepare()
        cs = ex.list_contacts(dec)
        log(f"E2E 联系人={len(cs)}")
        hits = [c for c in cs if (not kw) or kw in c["name"]
                or kw in c["nick"] or kw in c["wxid"]]
        if not hits:
            log("E2E FAIL: 没找到匹配联系人"); return 3
        target = hits[0]
        msgs, me = ex.collect(target["wxid"])
        if not msgs:
            log("E2E FAIL: 该会话无消息"); return 4
        doc = ex.write_word(msgs, target, me, outdir, "我")
        tp = ex.write_txt(msgs, target, me, outdir, "我")
        media = ex.gather_media(msgs, outdir, target_wxid=target["wxid"])
        with zipfile.ZipFile(doc) as z:
            xml_ok = "word/document.xml" in z.namelist()
        ex.cleanup()
        log(f"E2E OK msgs={len(msgs)} doc={doc} txt={tp} media={media} "
            f"docx_xml_ok={xml_ok}")
        return 0 if xml_ok else 5
    except Exception:
        log("E2E FAIL: " + traceback.format_exc())
        return 1


if len(sys.argv) > 1 and sys.argv[1] == "--e2e":
    sys.exit(_headless_e2e())


import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog

elog("tkinter 导入完成")

FONT = "Microsoft YaHei UI"

# ---------------------------------------------------------------- 配色
C_BG = "#F7F6F2"        # 米白窗口底色
C_CARD = "#FFFFFF"      # 卡片
C_LINE = "#E7E4DC"      # 分割线 / 边框
C_TEXT = "#23262B"      # 主文字
C_SUB = "#8B8F98"       # 次要文字
C_ACCENT = "#07C160"    # 微信绿
C_ACCENT_D = "#068F47"  # 主色按下
C_SOFT = "#E8F7EF"      # 选中底色
C_FRIEND = "#07C160"
C_STRANGER = "#F0932B"
C_MUTED = "#B7BBC3"     # 本机无聊天记录的联系人
C_DANGER = "#D9534F"
C_LOG_BG = "#1E2230"
C_LOG_FG = "#CDD6E4"


class Card(tk.Frame):
    """圆角卡片：底色画在 Canvas 上，内容用一个 Frame 叠在上面。

    tkinter 没有原生圆角，这里用 canvas.create_polygon 的 smooth 参数
    画圆角矩形，窗口缩放时重绘。
    """

    def __init__(self, master, padding=14, radius=14, bg=C_CARD,
                 stretch=False, fixed=0, **kw):
        super().__init__(master, bg=C_BG, highlightthickness=0, **kw)
        self._radius = radius
        self._fill = bg
        self._stretch = stretch
        self._fixed = fixed
        self._cv = tk.Canvas(self, bg=C_BG, highlightthickness=0, bd=0)
        self._cv.pack(fill="both", expand=True)
        self.body = tk.Frame(self._cv, bg=bg)
        self._pad = padding
        self._win = self._cv.create_window(padding, padding,
                                           window=self.body, anchor="nw")
        self._cv.bind("<Configure>", self._redraw)
        self.body.bind("<Configure>", self._sync_height)

    def _sync_height(self, _e=None):
        try:
            if self._fixed:                       # 固定高度的卡片（输出区）
                if int(self._cv.cget("height") or 0) != self._fixed:
                    self._cv.configure(height=self._fixed)
                return
            h = self.body.winfo_reqheight() + self._pad * 2
            if int(self._cv.cget("height") or 0) != h and not self._fixed:
                self._cv.configure(height=h)
        except Exception:
            pass

    _expanded = True

    def _redraw(self, _e=None):
        try:
            w = self._cv.winfo_width()
            h = self._cv.winfo_height()
            if w < 4 or h < 4:
                return
            r = min(self._radius, w // 2, h // 2)
            self._cv.delete("bgshape")
            self._cv.create_polygon(
                self._pts(w, h, r), smooth=True, splinesteps=24,
                fill=self._fill, outline=C_LINE, width=1, tags="bgshape")
            self._cv.tag_lower("bgshape")
            self._cv.itemconfigure(self._win, width=max(40, w - self._pad * 2))
            # 关键修复：body 必须撑满 Canvas 高度，否则卡片内容被 reqheight 掐死，
            # 里面 expand=True 的列表只分到很矮的一块，看着像「表格被截断」。
            if self._stretch:
                self._cv.itemconfigure(
                    self._win, height=max(40, h - self._pad * 2))
        except Exception:
            pass

    @staticmethod
    def _pts(w, h, r):
        return [r, 0, w - r, 0, w, 0, w, r, w, h - r, w, h,
                w - r, h, r, h, 0, h, 0, h - r, 0, r, 0, 0]


class RoundButton(tk.Canvas):
    """圆角按钮（tk 的 ttk 按钮没法做圆角，自绘一个）。"""

    def __init__(self, master, text, command, width=120, height=38,
                 radius=10, bg=C_ACCENT, fg="#FFFFFF", font=None,
                 hover=None, parent_bg=C_CARD, sub=False):
        self._bw, self._bh, self._br = width, height, radius
        self._bg = bg
        self._hover = hover or C_ACCENT_D
        self._fg = fg
        self._cmd = command
        self._enabled = True
        self._sub = sub
        super().__init__(master, width=width, height=height, bg=parent_bg,
                         highlightthickness=0, bd=0, cursor="hand2")
        self._font = font or tkfont.Font(family=FONT, size=10,
                                         weight="bold" if not sub else "normal")
        self._paint(bg)
        self._t = self.create_text(width // 2, height // 2, text=text,
                                   fill=fg, font=self._font)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _paint(self, color, outline=""):
        r = self._br
        w, h = self._bw, self._bh
        pts = [r, 0, w - r, 0, w, 0, w, r, w, h - r, w, h,
               w - r, h, r, h, 0, h, 0, h - r, 0, r, 0, 0]
        self.delete("shape")
        self.create_polygon(pts, smooth=True, splinesteps=24, fill=color,
                            outline=outline, tags="shape")
        self.tag_lower("shape")
        if getattr(self, "_t", None) is not None:
            self.tag_raise(self._t)

    def _on_enter(self, _e):
        if self._enabled:
            self._paint(self._hover)

    def _on_leave(self, _e):
        if self._enabled:
            self._paint(self._bg)

    def _on_click(self, _e):
        if self._enabled and self._cmd:
            self._cmd()

    def set_enabled(self, on: bool):
        self._enabled = bool(on)
        self.configure(cursor="hand2" if on else "arrow")
        self._paint(self._bg if on else "#D8DEE9")
        self.itemconfigure(self._t, fill=self._fg if on else "#A9B1C1")

    def set_text(self, s: str):
        self.itemconfigure(self._t, text=s)


def _apply_icon(win) -> None:
    """设置窗口 + 任务栏图标。找不到就跳过，绝不影响启动。"""
    cands: list[Path] = [BASE / "app.ico"]
    if getattr(sys, "frozen", False):
        cands.append(Path(sys.executable).parent / "app.ico")
        cands.append(Path(sys.executable).parent / "_internal" / "app.ico")
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        cands.append(Path(meipass) / "app.ico")
    for c in cands:
        try:
            if c.is_file():
                win.iconbitmap(default=str(c))
                log(f"图标已加载：{c}")
                return
        except Exception as e:  # noqa: BLE001
            log(f"图标加载失败 {c}: {e}")
    log("未找到 app.ico，使用默认图标")


class App(tk.Tk):
    def __init__(self):
        log("App.__init__ 开始")
        super().__init__()
        self.title(APP)
        _apply_icon(self)
        self.geometry("1120x800")
        self.minsize(1000, 660)
        self.configure(bg=C_BG)

        self.ex: core.Exporter | None = None
        self.dec: Path | None = None
        self.contacts: list[dict] = []
        self.shown: list[dict] = []
        self.filtered: list[dict] = []
        self.busy = False
        self.q: queue.Queue = queue.Queue()
        self.last_doc: str = ""
        self._cancel_evt = threading.Event()
        self._t0 = 0.0
        self.group = tk.StringVar(value="好友")
        self.search_var: tk.StringVar | None = None

        self._build()
        self.bind_all("<Control-f>", self._focus_search)
        log("UI 构建完成")
        self.after(100, self._pump)
        self.after(250, self._bootstrap)
        if getattr(sys, "frozen", False):
            threading.Thread(target=self._integrity, daemon=True).start()
        log("App.__init__ 结束")

    # ------------------------------------------------------------ UI
    def _build(self):
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(family=FONT, size=10)
        tkfont.nametofont("TkTextFont").configure(family=FONT, size=10)

        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("TLabelframe", background=C_CARD, borderwidth=0)
        st.configure("TLabelframe.Label", background=C_CARD, foreground=C_SUB,
                     font=(FONT, 9))
        st.configure("Treeview", background=C_CARD, fieldbackground=C_CARD,
                     foreground=C_TEXT, rowheight=34, borderwidth=0,
                     font=(FONT, 10))
        st.configure("Treeview.Heading", background=C_CARD, foreground=C_SUB,
                     relief="flat", font=(FONT, 9))
        st.map("Treeview", background=[("selected", C_SOFT)],
               foreground=[("selected", C_ACCENT_D)])
        st.configure("TProgressbar", troughcolor="#E9EDF4",
                     background=C_ACCENT, borderwidth=0, thickness=6)
        st.configure("TCheckbutton", background=C_CARD, foreground=C_TEXT,
                     font=(FONT, 10))
        st.map("TCheckbutton", background=[("active", C_CARD)])
        st.configure("TRadiobutton", background=C_CARD, foreground=C_TEXT,
                     font=(FONT, 10))
        st.map("TRadiobutton", background=[("active", C_CARD)])

        # ---- 顶部标题栏 ----
        head = tk.Frame(self, bg=C_BG)
        head.pack(fill="x", padx=24, pady=(16, 4))
        tk.Label(head, text="微信聊天记录导出", bg=C_BG, fg=C_TEXT,
                 font=(FONT, 16, "bold")).pack(side="left")
        tk.Label(head, text="选联系人 · 双击导出",
                 bg=C_BG, fg=C_SUB, font=(FONT, 9)).pack(side="left",
                                                        padx=(12, 0), pady=(6, 0))
        self.lbl_log = tk.Label(head, text="日志", bg=C_BG, fg=C_SUB,
                                font=(FONT, 9), cursor="hand2")
        self.lbl_log.pack(side="right", pady=(6, 0))
        self.lbl_log.bind("<Button-1>", lambda e: self._open_log())
        self.badge = tk.Label(head, text="● 检查权限…", bg=C_BG, fg=C_SUB,
                              font=(FONT, 9), cursor="hand2")
        self.badge.pack(side="right", padx=(0, 18), pady=(6, 0))
        self.badge.bind("<Button-1>", lambda e: self._elevate())

        # ---- 主体：左联系人 / 右操作面板 ----
        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True, padx=24, pady=(8, 0))

        # ---- 左列：联系人 ----
        c1 = Card(body, padding=14, stretch=True)
        c1.pack(side="left", fill="both", expand=True)
        b1 = c1.body
        b1.configure(bg=C_CARD)

        # 搜索（整行）
        self.search_var = tk.StringVar()
        self._ent = tk.Entry(b1, textvariable=self.search_var, width=24,
                             relief="flat", bg="#F5F4F0", fg="#A8A69F",
                             insertbackground=C_TEXT, font=(FONT, 10),
                             highlightthickness=1, highlightbackground=C_LINE,
                             highlightcolor=C_ACCENT)
        self._ent.pack(fill="x", ipady=6)
        self._ph = "搜索备注 / 昵称 / 微信号 / 拼音首字母（qrh）"
        self.search_var.set(self._ph)
        self._ent.bind("<FocusIn>", self._ph_in)
        self._ent.bind("<FocusOut>", self._ph_out)
        self._ent.bind("<Escape>", self._clear_search)
        self.search_var.trace_add("write", lambda *a: self._filter())

        # 分组切换 + 计数
        row = tk.Frame(b1, bg=C_CARD)
        row.pack(fill="x", pady=(8, 4))
        grp = tk.Frame(row, bg=C_CARD)
        grp.pack(side="left")
        self.btn_grp_f = self._pill(grp, "好友", "好友", self._switch_group)
        self.btn_grp_f.pack(side="left")
        self.btn_grp_s = self._pill(grp, "陌生人", "陌生人", self._switch_group)
        self.btn_grp_s.pack(side="left", padx=(6, 0))
        self.lbl_count = tk.Label(row, text="", bg=C_CARD, fg=C_SUB,
                                  font=(FONT, 9), anchor="e")
        self.lbl_count.pack(side="right")

        # 列表区：左字母索引 + 右表格
        mid = tk.Frame(b1, bg=C_CARD)
        mid.pack(fill="both", expand=True)

        self.idx = tk.Frame(mid, bg=C_CARD, width=30)
        self.idx.pack(side="left", fill="y", padx=(0, 8))
        self.idx.pack_propagate(False)
        self._build_index()

        wrap = tk.Frame(mid, bg=C_CARD, highlightthickness=1,
                        highlightbackground=C_LINE)
        wrap.pack(side="left", fill="both", expand=True)

        cols = ("mark", "name", "nick", "wxid")
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings", height=11)
        self.tree.heading("mark", text="")
        self.tree.column("mark", width=56, anchor="w", stretch=False)
        self.tree.heading("name", text="备注名 / 名称")
        self.tree.column("name", width=230, anchor="w")
        self.tree.heading("nick", text="昵称")
        self.tree.column("nick", width=210, anchor="w")
        self.tree.heading("wxid", text="微信号")
        self.tree.column("wxid", width=210, anchor="w")

        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        self.tree.tag_configure("stranger", foreground="#B26A00")
        self.tree.tag_configure("friend", foreground=C_TEXT)
        self.tree.tag_configure("nochat", foreground=C_MUTED)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._on_double)
        self.tree.bind("<Return>", lambda e: self._go())
        self.tree.bind("<Button-3>", self._on_right)
        self.tree.bind("<MouseWheel>",
                       lambda e: self.after(30, self._sync_hl), add="+")

        # ---- 右列：详情 / 选项 / 进度 ----
        right = tk.Frame(body, bg=C_BG, width=372)
        right.pack(side="left", fill="y", padx=(12, 0))
        right.pack_propagate(False)

        # 详情
        cd = Card(right, padding=14, fixed=132)
        cd.pack(fill="x", pady=(0, 10))
        bd = cd.body
        bd.configure(bg=C_CARD)
        self.lbl_name = tk.Label(bd, text="还没选联系人", bg=C_CARD, fg=C_TEXT,
                                 font=(FONT, 14, "bold"), anchor="w")
        self.lbl_name.pack(fill="x")
        self.lbl_meta = tk.Label(bd, text="在左边点一个联系人", bg=C_CARD,
                                 fg=C_SUB, font=(FONT, 9), anchor="w",
                                 justify="left")
        self.lbl_meta.pack(fill="x", pady=(4, 0))
        self.lbl_state = tk.Label(bd, text="", bg=C_CARD, fg=C_SUB,
                                  font=(FONT, 9), anchor="w", justify="left")
        self.lbl_state.pack(fill="x", pady=(2, 0))

        # 选项 + 主按钮
        co = Card(right, padding=14, fixed=196)
        co.pack(fill="x", pady=(0, 10))
        bo = co.body
        bo.configure(bg=C_CARD)
        tk.Label(bo, text="保存位置", bg=C_CARD, fg=C_SUB,
                 font=(FONT, 9)).pack(anchor="w")
        r2 = tk.Frame(bo, bg=C_CARD)
        r2.pack(fill="x", pady=(4, 8))
        self.out = tk.StringVar(
            value=str(Path.home() / "Documents" / "微信导出"))
        tk.Entry(r2, textvariable=self.out, relief="flat", bg="#F5F4F0",
                 fg=C_TEXT, font=(FONT, 9), highlightthickness=1,
                 highlightbackground=C_LINE, highlightcolor=C_ACCENT
                 ).pack(side="left", fill="x", expand=True, ipady=5)
        self.btn_browse = RoundButton(r2, "浏览", self._browse, width=64,
                                      height=32, radius=8, bg="#EDEBE5",
                                      fg=C_TEXT, hover="#E2DFD8", sub=True)
        self.btn_browse.pack(side="left", padx=(8, 0))

        r3 = tk.Frame(bo, bg=C_CARD)
        r3.pack(fill="x")
        self.do_txt = tk.BooleanVar(value=True)
        self.do_media = tk.BooleanVar(value=False)
        ttk.Checkbutton(r3, text="同时导出 TXT",
                        variable=self.do_txt).pack(side="left")
        ttk.Checkbutton(r3, text="归集视频 / 文件",
                        variable=self.do_media).pack(side="left", padx=(14, 0))

        rb = tk.Frame(bo, bg=C_CARD)
        rb.pack(fill="x", pady=(10, 0))
        self.btn_go = RoundButton(rb, "开始导出", self._go, width=150,
                                  height=42, radius=11, bg=C_ACCENT,
                                  fg="#FFFFFF", parent_bg=C_CARD,
                                  font=tkfont.Font(family=FONT, size=11,
                                                   weight="bold"))
        self.btn_go.pack(side="left")
        self.btn_go.set_enabled(False)
        self.btn_cancel = RoundButton(rb, "取消", self._cancel, width=86,
                                      height=42, radius=11,
                                      bg="#F4E7E5", fg=C_DANGER,
                                      hover="#EBD8D5", parent_bg=C_CARD,
                                      sub=True)
        self.btn_cancel.pack(side="left", padx=(10, 0))
        self.btn_cancel.set_enabled(False)

        # 进度 + 结果
        cp = Card(right, padding=14, fixed=178)
        cp.pack(fill="x")
        bp = cp.body
        bp.configure(bg=C_CARD)
        self.pb = ttk.Progressbar(bp, mode="determinate")
        self.pb.pack(fill="x", pady=(0, 8))
        self.prog_lbl = tk.Label(bp, text="等待开始", bg=C_CARD, fg=C_TEXT,
                                 font=(FONT, 9), anchor="w")
        self.prog_lbl.pack(fill="x")
        self.prog_eta = tk.Label(bp, text="", bg=C_CARD, fg=C_SUB,
                                 font=(FONT, 9), anchor="w")
        self.prog_eta.pack(fill="x", pady=(2, 0))
        r4 = tk.Frame(bp, bg=C_CARD)
        r4.pack(fill="x", pady=(10, 0))
        self.btn_open = RoundButton(r4, "打开文件夹", self._open,
                                    width=110, height=32, radius=8,
                                    bg="#EDEBE5", fg=C_TEXT,
                                    hover="#E2DFD8", parent_bg=C_CARD,
                                    sub=True)
        self.btn_open.pack(side="left")
        self.btn_open.set_enabled(False)
        self.result_lbl = tk.Label(r4, text="", bg=C_CARD, fg=C_ACCENT_D,
                                   font=(FONT, 9), anchor="w", justify="left")
        self.result_lbl.pack(side="left", fill="x", expand=True, padx=(10, 0))

        # ---- 底栏 ----
        bf = tk.Frame(self, bg=C_BG)
        bf.pack(fill="x", padx=24, pady=(10, 0))
        self.btn_reload = RoundButton(bf, "重新扫描", self._bootstrap,
                                      width=110, height=38, radius=10,
                                      bg="#EDEBE5", fg=C_TEXT,
                                      hover="#E2DFD8", parent_bg=C_BG,
                                      sub=True)
        self.btn_reload.pack(side="left")
        self.stat = tk.Label(bf, text="启动中 ...", bg=C_BG, fg=C_SUB,
                             font=(FONT, 9), anchor="e")
        self.stat.pack(side="right")

        # ---- 日志卡片 ----
        c3 = Card(self, padding=10)
        c3.pack(fill="x", padx=24, pady=(10, 14))
        b3 = c3.body
        b3.configure(bg=C_CARD)
        tk.Label(b3, text="运行日志", bg=C_CARD, fg=C_SUB,
                 font=(FONT, 9)).pack(anchor="w", pady=(0, 4))
        lw = tk.Frame(b3, bg=C_LOG_BG)
        lw.pack(fill="both", expand=True)
        self.txt = tk.Text(lw, height=5, bg=C_LOG_BG, fg=C_LOG_FG,
                           insertbackground="#fff", wrap="word",
                           relief="flat", bd=0, padx=10, pady=8,
                           font=("Consolas", 9))
        tsb = ttk.Scrollbar(lw, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=tsb.set, state="disabled")
        self.txt.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")

    def _build_index(self):
        self.idx_lbs = {}
        self.idx_on = set()
        self.cur_letter = ""
        for w in self.idx.winfo_children():
            w.destroy()
        for ch in ["#"] + [chr(c) for c in range(65, 91)]:
            lb = tk.Label(self.idx, text=ch, bg=C_CARD, fg="#CBD2DC",
                          font=(FONT, 8), cursor="arrow")
            lb.pack(fill="x")
            lb.bind("<Button-1>", lambda e, c=ch: self._jump(c))
            lb.bind("<Enter>", lambda e, w=lb, c=ch: self._idx_hover(w, c, True))
            lb.bind("<Leave>", lambda e, w=lb, c=ch: self._idx_hover(w, c, False))
            self.idx_lbs[ch] = lb

    def _paint_index(self):
        """可点的字母深一点，当前滚动到的字母加粗高亮。"""
        for ch, lb in self.idx_lbs.items():
            on = ch in self.idx_on
            cur = ch == self.cur_letter
            lb.configure(cursor="hand2" if on else "arrow",
                         fg=C_ACCENT_D if cur else (C_SUB if on else "#CBD2DC"),
                         font=(FONT, 8, "bold" if cur else "normal"))

    def _idx_hover(self, lb, ch, enter):
        if ch not in self.idx_on:
            return
        lb.configure(fg=C_ACCENT if enter or ch == self.cur_letter else C_SUB)

    def _pill(self, master, text, value, cmd):
        """分组切换按钮：选中态为浅蓝底。"""
        def paint(lbl, on):
            lbl.configure(bg=C_SOFT if on else "#F5F7FB",
                          fg=C_ACCENT_D if on else C_SUB)

        lbl = tk.Label(master, text=text, font=(FONT, 10),
                       padx=16, pady=6, cursor="hand2")
        paint(lbl, self.group.get() == value)
        lbl._value = value
        lbl._paint = paint
        lbl.bind("<Button-1>", lambda e: cmd(value))
        return lbl

    # ------------------------------------------------------------ 工具
    def _browse(self):
        d = filedialog.askdirectory(initialdir=self.out.get() or str(Path.home()))
        if d:
            self.out.set(d)

    def _open_log(self):
        if LOG_PATH and Path(LOG_PATH).exists():
            try:
                os.startfile(str(LOG_PATH))
                return
            except Exception:
                pass
        messagebox.showinfo(APP, f"日志：{LOG_PATH}")

    def _log(self, m: str):
        self.q.put(("log", m))
        log("UI| " + str(m))

    def _prog(self, c: int, t: int, m: str):
        self.q.put(("prog", (c, t, m)))

    def _pump(self):
        try:
            while True:
                k, v = self.q.get_nowait()
                if k == "log":
                    self.txt.configure(state="normal")
                    self.txt.insert("end", str(v) + "\n")
                    self.txt.see("end")
                    self.txt.configure(state="disabled")
                elif k == "status":
                    self.stat.configure(text=str(v))
                elif k == "prog":
                    c, t, m = v
                    if t > 0:
                        self.pb.stop()
                        self.pb.configure(mode="determinate",
                                          maximum=t, value=c)
                        self.prog_lbl.configure(text=m)
                        self.prog_eta.configure(text=self._eta(c, t))
                    else:
                        self.pb.configure(mode="indeterminate")
                        self.pb.start(12)
                        self.prog_lbl.configure(text=m)
                        self.prog_eta.configure(text="")
                elif k == "prog_stop":
                    self.pb.stop()
                    self.pb.configure(mode="determinate")
                    self.pb["value"] = 0
                    self.prog_eta.configure(text="")
                elif k == "contacts":
                    self.contacts = v
                    self._switch_group(self.group.get(), keep=True)
                    n_f = sum(1 for c in v if c["kind"] == "好友")
                    n_ok = sum(1 for c in v if c.get("has_msgs", True))
                    self.stat.configure(
                        text=f"就绪 · {len(v)} 个联系人"
                             f"（好友 {n_f}，可导出 {n_ok}）")
                    self.prog_lbl.configure(text="等待开始")
                    self._busy(False)
                    log(f"扫描完成，联系人 {len(v)} 个，可导出 {n_ok}")
                elif k == "fail":
                    self._busy(False)
                    self.q.put(("prog_stop", None))
                    self.stat.configure(text="出错")
                    self.prog_lbl.configure(text="出错了")
                    self._log("✗ " + str(v).splitlines()[0])
                    messagebox.showerror(APP, str(v)[:800])
                elif k == "cancelled":
                    self._busy(False)
                    self.q.put(("prog_stop", None))
                    self.stat.configure(text="已取消")
                    self.prog_lbl.configure(text="已取消，文件没有生成")
                    self._log("已取消")
                elif k == "done":
                    self._busy(False)
                    self.q.put(("prog_stop", None))
                    self.last_doc = v.get("doc", "")
                    self.stat.configure(text="导出完成")
                    self.prog_lbl.configure(text="导出完成 ✓")
                    self.btn_open.set_enabled(True)
                    doc = Path(self.last_doc)
                    size = (f"{doc.stat().st_size / 1024:.0f} KB"
                            if doc.exists() else "")
                    self.result_lbl.configure(text=f"{doc.name} · {size}")
                    self._log(f"完成：{self.last_doc}")
                    self._open()          # 自动打开并选中刚生成的文件
        except queue.Empty:
            pass
        except Exception:
            log("_pump 异常：" + traceback.format_exc())
        self.after(100, self._pump)

    def _eta(self, c: int, t: int) -> str:
        if not self._t0 or c <= 0 or t <= 0:
            return ""
        el = time.time() - self._t0
        left = el * (t - c) / c
        if left < 1:
            return "马上完成"
        if left < 60:
            return f"大约还要 {left:.0f} 秒"
        return f"大约还要 {left / 60:.1f} 分钟"

    def _busy(self, b: bool):
        self.busy = b
        self.btn_reload.set_enabled(not b)
        self.btn_cancel.set_enabled(bool(b) and not self._cancel_evt.is_set())
        if b:
            self.btn_go.set_enabled(False)
        else:
            c = self._cur()
            self.btn_go.set_enabled(bool(c and c.get("has_msgs", True)))

    # ------------------------------------------------------------ 分组 / 筛选
    def _switch_group(self, value, keep=False):
        self.group.set(value)
        for lbl in (self.btn_grp_f, self.btn_grp_s):
            lbl._paint(lbl, lbl._value == value)
        if not keep:
            self.tree.selection_remove(*self.tree.selection())
        self._filter()

    def _ph_in(self, _e):
        if self.search_var.get() == getattr(self, "_ph", ""):
            self.search_var.set("")
        self._ent.configure(fg=C_TEXT)

    def _ph_out(self, _e):
        if not self.search_var.get().strip():
            self.search_var.set(getattr(self, "_ph", ""))

    def _focus_search(self, _e=None):
        self._ent.focus_set()
        self._ent.select_range(0, "end")
        return "break"

    def _clear_search(self, _e=None):
        self.search_var.set(getattr(self, "_ph", ""))
        self._ent.configure(fg="#A8A69F")
        self._ent.tk_focusNext().focus_set()

    def _filter(self):
        raw = (self.search_var.get() if self.search_var else "").strip()
        if raw == getattr(self, "_ph", ""):
            raw = ""
        kw = raw.lower()
        base = [c for c in self.contacts if c["kind"] == self.group.get()]
        if kw:
            def hit(c: dict) -> bool:
                if any(kw in (c.get(k) or "").lower()
                       for k in ("name", "nick", "wxid", "alias")):
                    return True
                # 拼音首字母：打 qrh 就能找到「瞿瑞海」
                return (kw in core.pinyin_key(c["name"]).lower()
                        or kw in core.pinyin_key(c.get("nick") or "").lower())

            base = [c for c in base if hit(c)]
        self.idx_on = {c["letter"] for c in base}
        self.filtered = base
        self.shown = base[:1200]

        self.tree.delete(*self.tree.get_children())
        last_letter = None
        for i, c in enumerate(self.shown):
            lt = c["letter"]
            if lt != last_letter:
                self.tree.insert("", "end", iid=f"h{lt}",
                                 values=(lt, "", "", ""),
                                 tags=("__hdr__",))
                last_letter = lt
            if not c.get("has_msgs", True):
                tag = "nochat"
            else:
                tag = "stranger" if c["kind"] == "陌生人" else "friend"
            self.tree.insert("", "end", iid=str(i),
                             values=(lt, c["name"], c["nick"],
                                     c.get("alias") or c["wxid"]),
                             tags=(tag,))
        self.tree.tag_configure("__hdr__", background="#F7F6F2",
                                foreground=C_SUB, font=(FONT, 9, "bold"))
        n_ok = sum(1 for c in base if c.get("has_msgs", True))
        if not self.contacts:
            self.lbl_count.configure(
                text="还没有联系人：请先登录电脑版微信，再点「重新扫描」")
        else:
            self.lbl_count.configure(
                text=f"{self.group.get()} {len(base)} 个（可导出 {n_ok}）"
                     + (f" · 「{kw}」" if kw else ""))
        self._paint_index()
        self._sync_hl()
        self._detail(self._cur())

    def _jump(self, ch: str):
        """点字母跳到该分组，并显示这段有几个人。"""
        if ch not in self.idx_on:
            self.lbl_count.configure(text=f"{ch} 段没有联系人")
            return
        self.cur_letter = ch
        self._paint_index()
        iid = f"h{ch}"
        if self.tree.exists(iid):
            self.tree.see(iid)
            self.tree.selection_set(iid)
        n = sum(1 for c in self.shown if c["letter"] == ch)
        self.lbl_count.configure(text=f"{ch} 段 {n} 人")

    def _sync_hl(self):
        """滚动时同步字母索引的高亮。"""
        try:
            top = None
            for i in self.tree.get_children():
                b = self.tree.bbox(i)
                if b and b[1] >= 0:
                    top = i
                    break
            if top is None:
                return
            if top.startswith("h"):
                ch = top[1:]
            else:
                vals = self.tree.item(top, "values")
                ch = str(vals[0]) if vals else ""
            if ch and ch != self.cur_letter:
                self.cur_letter = ch
                self._paint_index()
        except Exception:
            pass

    def _cur(self) -> dict | None:
        s = self.tree.selection()
        if not s or str(s[0]).startswith("h"):
            return None
        try:
            return self.shown[int(s[0])]
        except (ValueError, IndexError):
            return None

    def _detail(self, c: dict | None):
        if not c:
            self.lbl_name.configure(text="还没选联系人")
            self.lbl_meta.configure(text="在左边点一个联系人")
            self.lbl_state.configure(text="", fg=C_SUB)
            self.btn_go.set_enabled(False)
            return
        self.lbl_name.configure(text=c["name"])
        meta = f"{c['kind']} · 微信号 {c.get('alias') or c['wxid']}"
        if c.get("nick") and c["nick"] != c["name"]:
            meta += f"\n昵称 {c['nick']}"
        self.lbl_meta.configure(text=meta)
        if c.get("has_msgs", True):
            self.lbl_state.configure(text="✓ 本机有聊天记录，可以导出",
                                     fg=C_ACCENT_D)
        else:
            self.lbl_state.configure(
                text="本机没有这个人的聊天记录，导不出内容", fg=C_STRANGER)
        self.btn_go.set_enabled(bool(c.get("has_msgs", True)) and not self.busy)

    def _on_select(self, _e=None):
        self._detail(self._cur())

    def _on_double(self, e):
        iid = self.tree.identify_row(e.y)
        if not iid or iid.startswith("h"):
            return
        self.tree.selection_set(iid)
        self._go()

    def _on_right(self, e):
        iid = self.tree.identify_row(e.y)
        if not iid or iid.startswith("h"):
            return
        self.tree.selection_set(iid)
        if not hasattr(self, "_menu"):
            m = tk.Menu(self, tearoff=0)
            m.add_command(label="导出聊天记录", command=self._go)
            m.add_command(label="复制微信号", command=self._copy_wxid)
            m.add_separator()
            m.add_command(label="打开导出文件夹", command=self._open)
            self._menu = m
        try:
            self._menu.tk_popup(e.x_root, e.y_root)
        finally:
            self._menu.grab_release()

    def _copy_wxid(self):
        c = self._cur()
        if not c:
            return
        txt = c.get("alias") or c["wxid"]
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.stat.configure(text=f"已复制微信号 {txt}")

    def _open(self):
        """有产物就定位到那个文件，没有就打开输出目录。"""
        try:
            if self.last_doc and Path(self.last_doc).exists():
                subprocess.Popen(["explorer", "/select,", str(self.last_doc)])
            else:
                os.startfile(str(Path(self.out.get())))
        except Exception as e:
            log(f"打开文件夹失败：{e}")
            messagebox.showinfo(APP, f"输出目录：{self.out.get()}")

    # ------------------------------------------------------------ 流程
    def _integrity(self):
        """打包文件自检：防拷贝不全/意外改动/顺手篡改；挡不住连清单一起改的人（那需要签名）。"""
        try:
            bad = core.verify_integrity(Path(sys.executable).parent)
            if bad:
                log(f"integrity 不一致: {bad[:5]}")
                self.q.put(("fail", "完整性自检不通过（可能被篡改或拷贝不全）：\n"
                                    + "\n".join(bad[:8])))
            else:
                log("integrity ok")
        except Exception as e:
            log(f"integrity 检查异常: {type(e).__name__}: {e}")

    def _bootstrap(self):
        elog("_bootstrap 进入")
        log("_bootstrap 进入")
        if self.busy:
            log("_bootstrap: busy，跳过")
            return

        is_adm = core.is_admin()
        elog(f"is_admin={is_adm}")
        log(f"is_admin() = {is_adm}")
        self._set_badge(is_adm)

        if not is_adm:
            # 打包版带 admin manifest，正常双击时不会走到这里；
            # 走到这里说明是源码运行或被降权，给一个可点的提示，不弹对话框。
            log("未提权，留在界面等用户点徽标")
            self._log("⚠ 当前不是管理员，读不到微信密钥。")
            self._log("   点右上角「未提权 · 点此提权」，或右键 exe →「以管理员身份运行」。")
            self.stat.configure(text="未提权 · 无法读取微信数据")
            self.prog_lbl.configure(text="未提权：点右上角提权")
            self.btn_go.set_enabled(False)
            return

        if not self._envcheck():
            return

        self._busy(True)
        self._log("")
        self._log("=" * 46)
        self.q.put(("prog", (0, 0, "正在提取密钥并解密数据库 ...")))

        def work():
            try:
                log("后台线程：开始 prepare()")
                self.ex = core.Exporter(self._log, self._prog)
                dec = self.ex.prepare()
                log(f"后台线程：prepare 完成 -> {dec}")
                self.dec = dec
                cs = self.ex.list_contacts(dec)
                log(f"后台线程：list_contacts 完成，{len(cs)} 个")
                self.q.put(("contacts", cs))
            except Exception as e:
                log("后台线程失败：" + traceback.format_exc())
                self.q.put(("fail", self._humanize(e)))

        threading.Thread(target=work, daemon=True).start()

    def _envcheck(self) -> bool:
        """启动自检：换到别的电脑时，直接说清哪一环不满足。"""
        try:
            issues = core.env_check()
        except Exception as e:
            self._log(f"! 自检本身出错（不影响使用）：{e}")
            return True
        for lv, msg in issues:
            self._log(("✗ " if lv == "err" else "! " if lv == "warn" else "· ") + msg)
        errs = [m for lv, m in issues if lv == "err"]
        if errs:
            self.stat.configure(text="环境自检未通过")
            messagebox.showerror(APP, "环境自检没通过：\n\n" + "\n\n".join(errs))
            return False
        return True

    def _set_badge(self, adm: bool):
        if adm:
            self.badge.configure(text="● 管理员权限", fg=C_ACCENT_D,
                                 cursor="arrow")
        else:
            self.badge.configure(text="● 未提权 · 点此提权", fg=C_DANGER,
                                 cursor="hand2")

    def _elevate(self):
        if core.is_admin():
            messagebox.showinfo(APP, "已经是管理员权限了。")
            return
        log("用户点了提权徽标")
        if core.relaunch_as_admin():
            self.destroy()
        else:
            messagebox.showerror(
                APP, "提权启动失败。请右键本程序 →「以管理员身份运行」。")

    def _humanize(self, e: Exception) -> str:
        """把底层异常翻成人话，界面上不出现 RuntimeError 这种字样。"""
        m = str(e)
        if isinstance(e, core.ExportCancelled):
            return "已取消。"
        if "微信没在运行" in m or "没在运行" in m:
            return "没找到电脑版微信。请先打开微信并登录，再点「重新扫描」。"
        if "没有任何消息" in m or "没有消息" in m:
            return "这个联系人在本机上没有聊天记录，换一个试试。"
        if "还没解密" in m or "解密目录" in m:
            return "数据还没准备好，点「重新扫描」重新来一次。"
        if isinstance(e, PermissionError) or "Access is denied" in m \
                or "权限" in m:
            return "没有权限读取微信的数据。点右上角「未提权 · 点此提权」重试。"
        return f"出错了：{m}"

    def _cancel(self):
        if not self.busy:
            return
        self._cancel_evt.set()
        self.btn_cancel.set_enabled(False)
        self.prog_lbl.configure(text="正在取消…")
        self._log("用户点了取消")

    def _go(self):
        target = self._cur()
        if not target:
            messagebox.showinfo(APP, "请先在左边选一个联系人")
            return
        if not target.get("has_msgs", True):
            messagebox.showinfo(
                APP, f"「{target['name']}」在本机上没有聊天记录，导不出内容。")
            return
        if not self.dec or not self.ex:
            messagebox.showinfo(APP, "还没扫描完，稍等一下")
            return
        if self.busy:
            return
        target = dict(target)
        log(f"开始导出：{target['name']} ({target['wxid']}) "
            f"kind={target['kind']}")
        outdir = Path(self.out.get())
        mynick = "我"
        want_txt = self.do_txt.get()
        want_media = self.do_media.get()

        self._cancel_evt.clear()
        self._t0 = time.time()
        self._busy(True)
        self.q.put(("prog", (0, 0, "正在读取消息 ...")))
        self._log("")
        self._log(f"导出：{target['name']}  [{target['kind']}]  {target['wxid']}")
        cancel = self._cancel_evt.is_set

        def work():
            try:
                msgs, me = self.ex.collect(target["wxid"],
                                           progress=self._prog,
                                           cancel=cancel)
                log(f"collect -> {len(msgs)} 条消息")
                if not msgs:
                    raise RuntimeError("这个会话在本机上没有任何消息记录")
                self._t0 = time.time()
                self.q.put(("prog", (0, 0, "正在写入 Word ...")))
                doc = self.ex.write_word(msgs, target, me, outdir, mynick,
                                         cancel=cancel)
                log(f"write_word -> {doc}")
                res = {"doc": str(doc)}
                if want_txt:
                    tp = self.ex.write_txt(msgs, target, me, outdir, mynick)
                    res["txt"] = str(tp)
                    log(f"write_txt -> {tp}")
                if want_media:
                    self._t0 = time.time()
                    self.q.put(("prog", (0, 0, "正在归集视频/文件 ...")))
                    res["media"] = self.ex.gather_media(
                        msgs, outdir, target_wxid=target["wxid"],
                        cancel=cancel)
                    log(f"gather_media -> {res['media']}")
                self.q.put(("done", res))
            except core.ExportCancelled:
                log("导出被用户取消")
                self.q.put(("cancelled", None))
            except Exception as e:
                log("导出失败：" + traceback.format_exc())
                self.q.put(("fail", self._humanize(e)))

        threading.Thread(target=work, daemon=True).start()

    def destroy(self):
        log("destroy() 调用，退出")
        try:
            if self.ex:
                self.ex.cleanup()
        except Exception:
            pass
        super().destroy()


def _excepthook(t, v, tb):
    log("未捕获异常：" + "".join(traceback.format_exception(t, v, tb)))
    try:
        messagebox.showerror(APP, f"程序出错：\n{t.__name__}: {v}\n\n日志：{LOG_PATH}")
    except Exception:
        pass


if __name__ == "__main__":
    elog("进入 __main__")
    sys.excepthook = _excepthook
    try:
        elog("开始构造 App()")
        a = App()
        elog("App() 构造完成，进 mainloop")
        a.protocol("WM_DELETE_WINDOW", a.destroy)
        a.mainloop()
        elog("mainloop 正常结束")
        log("mainloop 正常结束")
    except Exception:
        elog("主循环异常：" + traceback.format_exc())
        log("主循环异常：" + traceback.format_exc())
        raise