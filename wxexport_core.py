#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""微信聊天记录导出 - 核心库

把整个流程封装成可被 GUI 调用的类：
  1. 扫描微信数据库目录、从进程内存提取密钥、解密数据库
  2. 列出联系人与群聊
  3. 导出指定会话为 Word / TXT
  4. 把图片解密后内嵌进 Word、导出表情包
"""
from __future__ import annotations

import atexit
import concurrent.futures
import datetime
import functools
import hashlib
import json
import os
import re
import shutil
import sqlite3
import struct
import subprocess
import sys
import tempfile
import threading
import time
import zlib
import traceback
from pathlib import Path
from typing import Callable, Iterable

VENDOR = Path(__file__).resolve().parent / "vendor" / "wechat4-db-decrypt"
PAGE_SZ = 4096

# 打包成 GUI 程序（无控制台）后，每调一次 ffmpeg 就会新开一个控制台窗口；
# 并发导出图片时表现为「一堆黑窗一闪一闪地弹出来又消失」。
if os.name == "nt":
    _NO_WINDOW = 0x08000000  # CREATE_NO_WINDOW
    _NO_WINDOW_SI = subprocess.STARTUPINFO()
    _NO_WINDOW_SI.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    _NO_WINDOW_SI.wShowWindow = 0  # SW_HIDE
else:
    _NO_WINDOW = 0
    _NO_WINDOW_SI = None


class ExportCancelled(Exception):
    """用户在界面上点了「取消」。"""


# ------------------------------------------------------------------ 依赖
def find_python() -> str:
    """找一个装了本工具全部依赖的 python 解释器。

    打包成 exe 后 sys.executable 是 exe 自身，绝不能把它当 python 用
    （会变成「再启动一个自己」，而不是解释执行）。所以 frozen 时直接返回空。
    """
    import subprocess

    if getattr(__import__("sys"), "frozen", False):
        return ""

    def ok(py: str) -> bool:
        try:
            return subprocess.run([py, "-c", "import Crypto,docx,zstandard"],
                                  capture_output=True, timeout=60).returncode == 0
        except Exception:
            return False

    cands: list[str] = []
    env = os.environ.get("WX_EXPORT_PYTHON")
    if env:
        cands.append(env)
    for pat in (r"C:\Users\*\.dsh\dsh-runtimes\**\python.exe",
                r"C:\Users\*\AppData\Local\Programs\DeepSeek Harness\**\python.exe",
                r"C:\Users\*\AppData\Local\Programs\Python\**\python.exe"):
        try:
            cands.extend(str(p) for p in Path("C:/").glob(pat.lstrip("C:/")))
        except Exception:
            pass
    for n in ("python", "python3"):
        w = shutil.which(n)
        if w:
            cands.append(w)
    cands.append(os.sys.executable)
    seen = set()
    for c in cands:
        if not c or c in seen or not Path(c).exists():
            continue
        seen.add(c)
        if ok(c):
            return c
    return ""


def purge_stale_tmp(max_age_h: float = 0.0) -> int:
    """清掉 temp 里崩溃残留的解密明文目录（wxexport_*，超过 max_age_h 小时；默认 0 = 启动即全清）。

    GUI 有单实例锁，不会误删并发 GUI；脚本里自己并行跑多个 Exporter 的注意会互删。
    """
    import time
    root = Path(tempfile.gettempdir())
    cut = time.time() - max_age_h * 3600
    n = 0
    for d in root.glob("wxexport_*"):
        try:
            if d.is_dir() and d.stat().st_mtime < cut:
                shutil.rmtree(d, ignore_errors=True)
                n += 1
        except OSError:
            continue
    return n


def write_integrity_manifest(root: Path) -> Path:
    """给打包目录生成 integrity.json（程序自身文件的 sha256）。

    只收根目录下的文件 + 整个 _internal\\：用户往程序目录里放的其它文件夹
    （截图、导出结果等）不参与校验，否则一改就误报「被改动」。
    """
    files = {}
    targets = [p for p in root.iterdir() if p.is_file()]
    sub = root / "_internal"
    if sub.is_dir():
        targets += list(sub.rglob("*"))
    for p in sorted(targets):
        # log.txt 是运行时日志，每次启动都会改，不能进清单
        if p.is_file() and p.name not in ("integrity.json", "log.txt") \
                and ".bak" not in p.name:
            rel = str(p.relative_to(root)).replace("\\", "/")
            files[rel] = hashlib.sha256(p.read_bytes()).hexdigest()
    m = root / "integrity.json"
    m.write_text(json.dumps({"sha256": files}, ensure_ascii=False),
                 encoding="utf-8")
    return m


def verify_integrity(root: Path) -> list[str]:
    """校验打包目录与 integrity.json 是否一致，返回不一致清单（空 = 通过）。

    只能防拷贝不全 / 意外改动 / 顺手篡改；挡不住连清单一起改的人——那需要代码签名。
    """
    m = root / "integrity.json"
    if not m.exists():
        return ["缺 integrity.json"]
    try:
        want = json.loads(m.read_text(encoding="utf-8"))["sha256"]
    except Exception as e:
        return [f"integrity.json 读不了: {e}"]
    bad = []
    for rel, h in want.items():
        p = root / rel
        if not p.is_file():
            bad.append(f"缺文件 {rel}")
            continue
        try:
            if hashlib.sha256(p.read_bytes()).hexdigest() != h:
                bad.append(f"被改动 {rel}")
        except OSError as e:
            bad.append(f"读不了 {rel}: {e}")
    return bad


def missing_deps() -> list[str]:
    out = []
    for m in ("Crypto", "docx", "zstandard"):
        try:
            __import__(m)
        except ImportError:
            out.append(m)
    return out


def is_admin() -> bool:
    import ctypes
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> bool:
    """用管理员权限重新拉起自己（UAC 弹窗）。

    打包成 exe 后 sys.executable 就是 exe 自身，而且它已经内嵌了脚本，
    不能再把 sys.argv[0] 当脚本路径传过去——那样新进程会去打开一个
    并不存在的 main.py，表现为「点了没反应」。所以分两种情况：
      - frozen：只传 exe 本身（外加真实业务参数，排除 argv[0]）
      - 源码：传 [脚本, *argv[1:]]
    """
    import ctypes
    import sys
    try:
        if getattr(sys, "frozen", False):
            exe = sys.executable
            args = sys.argv[1:]
        else:
            exe = sys.executable
            args = [str(Path(sys.argv[0]).resolve()), *sys.argv[1:]]
        # 用 list2cmdline 正确处理含空格的路径
        params = subprocess.list2cmdline(args)
        workdir = str(Path(exe).parent)
        rc = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", exe, params or None, workdir, 1)
        return rc > 32
    except Exception:
        return False


# ------------------------------------------------------------------ 环境探测
def env_check() -> list[tuple[str, str]]:
    """换机器前的自检。返回 [(级别, 人话)]，级别为 ok / warn / err。

    目的：在别的电脑上出错时，用户能一眼看出是哪一环不满足，
    而不是只看到一句「解密失败」。
    """
    import platform
    import sys
    out: list[tuple[str, str]] = []

    # 1. 系统与架构
    try:
        arch = platform.machine()
        ver = sys.getwindowsversion()
        out.append(("ok", f"系统：Windows {ver.major}.{ver.minor} · {arch}"))
        if arch.lower() not in ("amd64", "x86_64"):
            out.append(("err", f"当前 CPU 架构是 {arch}，本程序只有 64 位版"
                               "（32 位 / ARM 电脑跑不了）。"))
    except Exception as e:
        out.append(("warn", f"系统信息读不出来：{e}"))

    # 2. 管理员
    if is_admin():
        out.append(("ok", "权限：管理员 ✓"))
    else:
        out.append(("err", "权限：不是管理员，读不到微信密钥。"
                           "请右键 exe →「以管理员身份运行」。"))

    # 3. 微信目录（区分 4.x / 3.x）
    db = find_db_storage()
    if db:
        kind = "微信 4.x" if "xwechat_files" in str(db) else "微信 3.x"
        out.append(("ok", f"微信目录：{kind} → {db}"))
    else:
        out.append(("err", "没找到微信数据目录（应该在 "
                           r"「文档\xwechat_files」或「文档\WeChat Files」）。"
                           "请先在电脑上登录一次微信。"))

    # 4. 微信是否在运行
    try:
        if wechat_running():
            out.append(("ok", "微信：正在运行 ✓"))
        else:
            out.append(("warn", "微信没在运行：请登录后点「重新扫描」。"))
    except Exception:
        out.append(("warn", "微信进程状态读不出来。"))

    # 5. 输出目录可写
    try:
        od = Path.home() / "Documents" / "微信导出"
        od.mkdir(parents=True, exist_ok=True)
        t = od / ".wtest"
        t.write_text("ok", encoding="utf-8")
        t.unlink()
        out.append(("ok", f"输出目录可写：{od}"))
    except Exception as e:
        out.append(("err", f"输出目录写不了：{e}"))

    return out


def find_db_storage() -> Path | None:
    roots = [Path(os.path.expanduser(r"~\Documents\xwechat_files")),
             Path(os.path.expanduser(r"~\Documents\WeChat Files")),
             Path(os.path.expanduser(r"~\AppData\Roaming\Tencent\xwechat_files"))]
    best = None
    for root in roots:
        if not root.is_dir():
            continue
        try:
            items = list(root.iterdir())
        except OSError:
            continue
        for d in items:
            ds = d / "db_storage"
            if ds.is_dir():
                try:
                    mt = ds.stat().st_mtime
                except OSError:
                    continue
                if best is None or mt > best[0]:
                    best = (mt, ds)
    return best[1] if best else None


def detect_version() -> str:
    vers = []
    for c in (Path(r"C:\Program Files\Tencent\Weixin"),
              Path(r"C:\Program Files (x86)\Tencent\Weixin"),
              Path(r"C:\Program Files\Tencent\WeChat"),
              Path(r"C:\Program Files (x86)\Tencent\WeChat")):
        if not c.is_dir():
            continue
        try:
            for d in c.iterdir():
                if d.is_dir() and re.fullmatch(r"\d+\.\d+\.\d+\.\d+", d.name):
                    vers.append(d.name)
        except OSError:
            continue
    if not vers:
        return ""
    vers.sort(key=lambda v: [int(x) for x in v.split(".")])
    return vers[-1]


def wechat_running() -> bool:
    try:
        import sys
        sys.path.insert(0, str(VENDOR / "src"))
        from wxexport.win import memory
        return bool(memory.list_weixin_processes())
    except Exception:
        return False


def open_db(p: Path) -> sqlite3.Connection:
    con = sqlite3.connect(f"file:{p}?mode=ro", uri=True)
    con.text_factory = lambda b: b.decode("utf-8", "replace")
    return con


# ------------------------------------------------------------------ 消息解析
TYPE_NAMES = {1: "文本", 3: "图片", 34: "语音", 42: "名片", 43: "视频",
              47: "表情", 48: "位置", 49: "应用消息", 50: "通话",
              10000: "系统", 10002: "撤回", 41: "文件", 44: "视频",
              62: "短视频", 66: "已删除"}
APPSUB = {"57": "引用回复", "6": "文件", "5": "链接", "33": "小程序",
          "36": "小程序", "19": "聊天记录", "2000": "转账", "2001": "红包",
          "3": "音乐", "4": "视频号", "63": "视频号直播", "51": "视频号"}


def _zstd(raw: bytes) -> str | None:
    try:
        import zstandard
        d = zstandard.ZstdDecompressor()
        try:
            return d.decompress(raw, max_output_size=1 << 22).decode(
                "utf-8", "replace")
        except Exception:
            with d.stream_reader(raw) as r:
                # 兜底也限 4MB，防对方构造压缩炸弹消息撑爆内存
                return r.read(1 << 22).decode("utf-8", "replace")
    except Exception:
        return None


def _gb_initial(ch: str) -> str | None:
    """用 GB2312 一级字库（按拼音排序）取汉字拼音首字母。"""
    try:
        gb = ch.encode("gb2312")
    except Exception:
        return None
    if len(gb) != 2:
        return None
    v = (gb[0] << 8) | gb[1]
    if not (0xB0A1 <= v <= 0xD7F9):
        return "#"
    for code, letter in reversed(_GB_SEG):
        if v >= code:
            return letter
    return "#"


# GB2312 一级汉字 23 段的段首（实测解码：啊 芭 擦 搭 蛾 发 噶 哈 击 喀 垃 妈
# 拿 哦 啪 期 然 撒 塌 挖 昔 压 匝）
_GB_SEG = (
    (0xB0A1, "A"), (0xB0C5, "B"), (0xB2C1, "C"), (0xB4EE, "D"),
    (0xB6EA, "E"), (0xB7A2, "F"), (0xB8C1, "G"), (0xB9FE, "H"),
    (0xBBF7, "J"), (0xBFA6, "K"), (0xC0AC, "L"), (0xC2E8, "M"),
    (0xC4C3, "N"), (0xC5B6, "O"), (0xC5BE, "P"), (0xC6DA, "Q"),
    (0xC8BB, "R"), (0xC8F6, "S"), (0xCBFA, "T"), (0xCDDA, "W"),
    (0xCEF4, "X"), (0xD1B9, "Y"), (0xD4D1, "Z"),
)


# Windows 中文排序（zh-CN）的 23 个段首，一级字库按拼音排列，二级字也能排进去
_BOUND = (
    ("啊", "A"), ("芭", "B"), ("擦", "C"), ("搭", "D"), ("蛾", "E"),
    ("发", "F"), ("噶", "G"), ("哈", "H"), ("击", "J"), ("喀", "K"),
    ("垃", "L"), ("妈", "M"), ("拿", "N"), ("哦", "O"), ("啪", "P"),
    ("期", "Q"), ("然", "R"), ("撒", "S"), ("塌", "T"), ("挖", "W"),
    ("昔", "X"), ("压", "Y"), ("匝", "Z"),
)


def _coll_initial(ch: str) -> str | None:
    """用 Windows 自带的中文排序（zh-CN 按拼音）判断首字母。

    GB2312 一级字库只覆盖 3755 个常用字，「瞿 / 覃 / 冼」这类二级字
    用老办法只能落到 '#'。Windows 的 zh-CN 排序本身按拼音，
    拿它和 23 个段首逐个比就能把任意汉字归位。
    返回 None 表示这套机制不可用，调用方退回 GB 表。
    """
    try:
        import ctypes
        from ctypes import wintypes
        f = ctypes.windll.kernel32.CompareStringEx
        f.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.LPCWSTR,
                      ctypes.c_int, wintypes.LPCWSTR, ctypes.c_int,
                      ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t]
        f.restype = ctypes.c_int
        for ref, letter in reversed(_BOUND):
            r = f("zh-CN", 0, ch, 1, ref, 1, None, None, 0)
            if r == 0:
                return None          # 排序调用失败
            if r > 2:                # CSTR_GREATER_THAN
                return letter
        return "#"                   # 比「啊」还小：符号类
    except Exception:
        return None


@functools.lru_cache(maxsize=8192)
def _initial(ch: str) -> str | None:
    """单字 → 拼音首字母，优先用系统排序，失败再退回 GB2312 表。"""
    r = _coll_initial(ch)
    return _gb_initial(ch) if r is None else r


def pinyin_key(s: str) -> str:
    """把显示名转成可排序的键：汉字取拼音首字母，ASCII 大写保留，其余忽略。"""
    if not s:
        return ""
    out: list[str] = []
    for ch in s:
        if ord(ch) < 128:
            if ch.isalnum():
                out.append(ch.upper())
            continue
        ini = _initial(ch)
        if ini:
            out.append(ini)
        # emoji / 符号等无法判断音序的字符直接跳过
    return "".join(out)


def pinyin_initial(s: str) -> str:
    """返回显示名首字（跳过符号/emoji）的拼音首字母，判断不出返回 '#'。"""
    for ch in s or "":
        if ord(ch) < 128:
            return ch.upper() if ch.isalnum() else "#"
        ini = _initial(ch)
        if ini:
            return ini
    return "#"


def letter_of(name: str) -> str:
    """通讯录分组字母：A-Z，非字母开头归到 '#'。"""
    c0 = pinyin_initial(name)
    if not c0:
        return "#"
    c0 = c0[0]
    if c0.isascii() and c0.isalpha():
        return c0
    return "#"


def norm_content(content, mtype: int, rename: dict[str, str] | None = None) -> str:
    rename = rename or {}
    if isinstance(content, bytes) and content[:4] == bytes.fromhex("28b52ffd"):
        d = _zstd(content)
        if d is not None:
            content = d
    s = (content.decode("utf-8", "replace") if isinstance(content, bytes)
         else (content or ""))
    s = "".join(ch for ch in s if ch in "\n\t" or ord(ch) >= 32).strip()
    for a, b in rename.items():
        if a and a != b:
            s = s.replace(a, b)
    t = mtype
    if t == 1:
        return s or "（空）"
    if t == 3:
        return "[图片]"
    if t == 34:
        return "[语音]"
    if t in (43, 44, 62):
        return "[视频]"
    if t == 47:
        return "[表情]"
    if t == 41:
        m = re.search(r"<title>(.*?)</title>", s, re.S)
        return f"[文件] {m.group(1).strip()}" if m else "[文件]"
    if t == 49:
        m = re.search(r"<title>(.*?)</title>", s, re.S)
        if m:
            return f"[链接] {m.group(1).strip()}"
        m = re.search(r"<des>(.*?)</des>", s, re.S)
        if m:
            return f"[分享] {m.group(1).strip()[:80]}"
        return "[应用消息]"
    if t == 10000:
        m = re.search(r"<content>(.*?)</content>", s, re.S)
        if m:
            return m.group(1).strip()
        m = re.search(r"<replacemsg><!\[CDATA\[(.*?)\]\]></replacemsg>",
                      s, re.S)
        if m:
            return m.group(1).strip()
        return "（系统消息）"
    if t == 10002:
        return "[撤回了一条消息]"
    if t == 50:
        return "[通话]"
    if "<?xml" in s or "<msg>" in s or "<appmsg" in s:
        title = des = ""
        m = re.search(r"<title>(.*?)</title>", s, re.S)
        if m:
            title = m.group(1).strip()
        m2 = re.search(r"<des>(.*?)</des>", s, re.S)
        if m2:
            des = m2.group(1).strip()
        m3 = re.search(r"<type>(\d+)</type>", s, re.S)
        kind = APPSUB.get(m3.group(1) if m3 else "", "应用消息")
        parts = [x for x in (title, des) if x]
        return f"[{kind}] " + (" / ".join(parts)[:200] if parts else "（内容已省略）")
    tn = TYPE_NAMES.get(t, f"type{t}")
    return f"[{tn}] {s[:300]}" if s else f"[{tn}]"


# ------------------------------------------------------------------ 主流程
class Exporter:
    def __init__(self, log: Callable[[str], None] | None = None,
                 progress: Callable[[int, int, str], None] | None = None):
        self.log = log or (lambda m: None)
        self.progress = progress or (lambda c, t, m: None)
        self.tmp: Path | None = None
        self.dec_dir: Path | None = None
        self.keys: dict[str, str] = {}
        self.db_dir: Path | None = None
        self.contacts: list[dict] = []
        self.image_uin: int | None = None   # 图片/表情包密钥共用的 seed

    # ---- 1. 密钥 + 解密 ----
    def prepare(self) -> Path:
        import sys
        sys.path.insert(0, str(VENDOR / "src"))
        from wxexport.win import keyscan, memory
        from wxexport.decrypt import decrypt_database, merge_wal

        if not memory.is_admin():
            raise PermissionError(
                "需要管理员权限才能读取微信进程内存中的数据库密钥。\n"
                "请用「以管理员身份运行」重新打开本程序。")

        db = find_db_storage()
        if not db:
            raise RuntimeError(
                "没找到微信数据库目录。\n"
                "预期位置：Documents\\xwechat_files\\<账号>\\db_storage")
        self.db_dir = db
        self.log(f"数据库目录：{db}")

        ver = detect_version()
        self.log(f"微信版本：{ver or '未知'}")

        if not memory.list_weixin_processes():
            raise RuntimeError("微信没在运行。\n请先打开并登录电脑版微信，再点「重新扫描」。")

        self.progress(0, 0, "正在扫描微信内存提取密钥（约 10-60 秒）...")
        result = keyscan.hunt(db, version=ver,
                              strategies="config,legacy,passphrase",
                              deep=True, log=self.log)
        keys = {rel: info["enc_key"] for rel, info in result.mapping.items()}
        self.keys = keys
        self.log(f"密钥获取：{len(keys)}/{len(result.inventory.databases)}")
        if not keys:
            raise RuntimeError(
                "没能提取到任何数据库密钥。\n\n"
                "常见原因：\n"
                "1) 不是管理员权限\n"
                "2) 微信未登录（只到扫码界面）\n"
                "3) 微信版本太新，内存里没有明文密钥\n\n"
                "请先登录微信主页（能看到聊天列表），再重试。")

        if self.tmp and self.tmp.exists():
            shutil.rmtree(self.tmp, ignore_errors=True)
        purge_stale_tmp()
        self.tmp = Path(tempfile.mkdtemp(prefix="wxexport_"))
        atexit.register(self.cleanup)   # 异常退出时也尽量收掉明文目录
        dec = self.tmp / "dec"
        dec.mkdir()

        self.progress(0, 0, "正在解密数据库 ...")
        done = 0
        failed = 0
        for rel, hk in sorted(keys.items()):
            src = db / rel
            if not src.exists():
                continue
            dst = dec / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            try:
                key = bytes.fromhex(hk)
                decrypt_database(str(src), str(dst), key)
                try:
                    merge_wal(str(src), str(dst), key)
                except Exception:
                    pass
                done += 1
            except Exception as e:
                failed += 1
                self.log(f"  [跳过] {rel}: {type(e).__name__}")
        self.log(f"解密完成：成功 {done} 个，失败 {failed} 个")
        if done == 0:
            raise RuntimeError("所有数据库解密都失败了。")
        self.dec_dir = dec
        return dec

    # ---- 2. 联系人（仅一对一：好友 / 陌生人，排除群聊与公众号）----
    def list_contacts(self, dec: Path | None = None) -> list[dict]:
        """只返回能导出聊天记录的一对一联系人。

        微信 4.x 的 contact 表里：
          local_type = 1 → 好友（双向）
          local_type = 3 → 陌生人（含已删除好友，本地仍保留记录）
          local_type = 2 → 群聊（排除）
          local_type = 0/4/5/6 → 系统号 / 公众号 / 服务号（排除）
        """
        dec = dec or self.dec_dir
        if not dec:
            raise RuntimeError("还没解密")
        out: list[dict] = []
        # 先扫一遍 message/*.db，收集所有真实存在的会话表名。
        # 只列「有会话表」的联系人：contact 表会长期保留已删除好友、
        # 群友、名片往来等本地残留，它们没有会话记录，也导不出内容。
        live: set[str] = set()
        msgdir = dec / "message"
        if msgdir.is_dir():
            for mdb in sorted(msgdir.glob("message_*.db")):
                try:
                    c2 = open_db(mdb)
                    try:
                        for (t,) in c2.execute(
                                "SELECT name FROM sqlite_master "
                                "WHERE type='table' AND name LIKE 'Msg_%'"):
                            live.add(t)
                    finally:
                        c2.close()
                except Exception:
                    continue
        self.log(f"本地会话表：{len(live)} 个")
        db = dec / "contact" / "contact.db"
        if db.exists():
            con = open_db(db)
            try:
                cols = {r[1] for r in con.execute("PRAGMA table_info(contact)")}
                nick = "nick_name" if "nick_name" in cols else "''"
                rmk = "remark" if "remark" in cols else "''"
                als = "alias" if "alias" in cols else "''"
                ltp = "local_type" if "local_type" in cols else "0"
                rows = con.execute(
                    f"SELECT username, COALESCE(NULLIF({rmk},''),"
                    f"NULLIF({nick},''),username), COALESCE({nick},''),"
                    f"COALESCE({als},''), COALESCE({ltp},0) FROM contact").fetchall()
            finally:
                con.close()

            for un, disp, nk, al, lt in rows:
                if not un:
                    continue
                # 公众号 / 服务号（gh_ 开头）不是真人：微信把它们记成
                # local_type=1/3，混在好友和陌生人里，列表里必须剔掉。
                if un.startswith("gh_"):
                    continue
                name = (disp or un).strip() or un
                rec = {"wxid": un, "name": name, "nick": (nk or "").strip(),
                       "alias": (al or "").strip(),
                       "kind": "", "sort_key": "", "letter": "#"}
                if lt == 1:
                    rec["kind"] = "好友"
                elif lt == 3:
                    rec["kind"] = "陌生人"
                else:
                    continue            # 群聊 / 公众号 / 系统号一律不收
                # 没有会话表的联系人本地没有任何聊天记录：照样列出来，
                # 但标记 has_msgs=False，界面灰显并说明「本机无聊天记录」，
                # 免得用户以为好友丢了（以前是直接过滤掉）。
                rec["has_msgs"] = (
                    "Msg_" + hashlib.md5(un.encode()).hexdigest()) in live
                if not rec["name"]:
                    continue
                rec["sort_key"] = pinyin_key(rec["name"])
                rec["letter"] = letter_of(rec["name"])
                out.append(rec)
        # 按拼音键排序，'#' 段排最后
        out.sort(key=lambda x: (x["letter"] == "#", x["sort_key"] or "z" * 9,
                                x["name"]))
        self.contacts = out
        return out

    # ---- 3. 拉消息 ----
    def collect(self, target_wxid: str, dec: Path | None = None,
                progress: Callable[[int, int, str], None] | None = None,
                cancel: Callable[[], bool] | None = None):
        dec = dec or self.dec_dir
        if not dec:
            raise RuntimeError("还没解密")
        progress = progress or (lambda c, t, m: None)
        tbl = "Msg_" + hashlib.md5(target_wxid.encode()).hexdigest()
        msgs: list[dict] = []
        me = ""
        msgdir = dec / "message"
        if not msgdir.is_dir():
            raise RuntimeError("解密目录里没有 message 子目录")

        # 先数一遍总量，进度条才能按真实条数走（分库存储，一个会话可能在多个库里）
        plans: list[tuple[Path, int]] = []
        total = 0
        for db in sorted(msgdir.glob("message_*.db")):
            con = open_db(db)
            try:
                names = {r[0] for r in con.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'")}
                if tbl not in names:
                    continue
                n = con.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
                plans.append((db, n))
                total += n
            finally:
                con.close()
        if not plans:
            self.log("这个会话没有对应的消息表")
            return msgs, me
        progress(0, total, f"正在读取消息 0/{total}")

        done = 0
        for db, _n in plans:
            con = open_db(db)
            try:
                smap = {rid: un for rid, un in
                        con.execute("SELECT rowid, user_name FROM Name2Id")}
                for lid, ltype, sid, ct, content in con.execute(
                        f"SELECT local_id, local_type, real_sender_id,"
                        f" create_time, message_content FROM {tbl}"):
                    if cancel and cancel():
                        raise ExportCancelled("已取消")
                    msgs.append({"type": ltype, "sender": smap.get(sid, ""),
                                 "time": ct, "local_id": lid,
                                 "content": content})
                    done += 1
                    if done % 200 == 0 or done == total:
                        progress(done, total,
                                 f"正在读取消息 {done}/{total}")
                # 先认登录账号自己：一个会话里不止两个 wxid 时（被拉进来的陌生人、
                # 旧账号 id 等），靠「唯一其它 wxid」会认错，文档里就变成一串裸 wxid。
                acc = self._account_wxid()
                if not me and acc and acc in set(smap.values()):
                    me = acc
                others = {u for u in smap.values()
                          if u and u != target_wxid and u.startswith("wxid_")}
                if not me:
                    if len(others) == 1:
                        me = next(iter(others))
                    elif others:
                        cdb = dec / "contact" / "contact.db"
                        if cdb.exists():
                            c2 = open_db(cdb)
                            try:
                                row = c2.execute(
                                    "SELECT username FROM contact WHERE username IN (%s)"
                                    % ",".join("?" * len(others)),
                                    tuple(others)).fetchone()
                                if row:
                                    me = row[0]
                            finally:
                                c2.close()
            finally:
                con.close()
        msgs.sort(key=lambda m: (m["time"], m.get("local_id", 0)))
        self.log(f"收集到 {len(msgs)} 条消息，我方 wxid={me or '未知'}")
        return msgs, me

    # ---- 4. 写 Word ----
    def _add_picture(self, doc, data: bytes, max_w: float = 4.2,
                     max_h: float = 6.5) -> None:
        """按比例塞进页面：宽 ≤ max_w 英寸、高 ≤ max_h 英寸。"""
        from docx.shared import Inches
        try:
            import io
            from PIL import Image
            with Image.open(io.BytesIO(data)) as im:
                w, h = im.size
            win, hin = w / 96.0, h / 96.0
            k = min(max_w / win, max_h / hin, 1.0) if win and hin else 1.0
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25)
            p.add_run().add_picture(io.BytesIO(data), width=Inches(win * k))
        except Exception:
            self.log("图片嵌入失败：" + traceback.format_exc(limit=1))


    def write_word(self, msgs: list[dict], target: dict, me_wxid: str,
                   outdir: Path, my_nick: str = "我",
                   rename: dict[str, str] | None = None,
                   cancel: Callable[[], bool] | None = None) -> Path:
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor

        rename = rename or {}
        if not msgs:
            raise RuntimeError("这个会话没有任何消息")

        dname = target["name"]
        twxid = target["wxid"]
        t0 = datetime.datetime.fromtimestamp(msgs[0]["time"])
        t1 = datetime.datetime.fromtimestamp(msgs[-1]["time"])
        mine = sum(1 for m in msgs if m["sender"] == me_wxid)
        theirs = sum(1 for m in msgs if m["sender"] == twxid)

        doc = Document()
        doc.add_heading(f"与 {dname} 的聊天记录", 0)
        meta = doc.add_paragraph()
        meta.add_run(f"备注：{dname}").bold = True
        if target.get("nick") and target["nick"] != dname:
            meta.add_run(f"（昵称 {target['nick']}）")
        meta.add_run(f"\n消息总数：{len(msgs)} 条（我 {mine} / 对方 {theirs}）\n")
        meta.add_run(f"时间范围：{t0:%Y-%m-%d %H:%M} ~ {t1:%Y-%m-%d %H:%M}\n")
        meta.add_run(f"导出时间：{datetime.datetime.now():%Y-%m-%d %H:%M}")
        doc.add_paragraph()

        last_day = None
        total = len(msgs)
        for i, m in enumerate(msgs):
            if cancel and cancel():
                raise ExportCancelled("已取消")
            dt = datetime.datetime.fromtimestamp(m["time"])
            if dt.date() != last_day:
                doc.add_heading(f"{dt:%Y年%m月%d日}", level=2)
                last_day = dt.date()
            if m["sender"] == me_wxid:
                who, color = my_nick, RGBColor(0x1A, 0x73, 0xE8)
            elif m["sender"] == twxid:
                who, color = dname, RGBColor(0xC0, 0x39, 0x2B)
            else:
                who, color = (m["sender"] or "系统"), RGBColor(0x88, 0x88, 0x88)
            p = doc.add_paragraph()
            ts = p.add_run(f"[{dt:%H:%M:%S}] ")
            ts.font.size = Pt(8)
            ts.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
            nm = p.add_run(f"{who}：")
            nm.bold = True
            nm.font.size = Pt(9)
            nm.font.color.rgb = color
            img = m.get("img")
            sticker = bool(img) and m.get("img_kind") == "sticker"
            if not sticker:      # 表情已经画出来了，就不再写一行「[表情]」
                bd = p.add_run(norm_content(m["content"], m["type"], rename))
                bd.font.size = Pt(10)
            if img:
                if sticker:
                    self._add_picture(doc, img[0], max_w=1.5, max_h=1.5)
                else:
                    self._add_picture(doc, img[0])
            if i % 200 == 0:
                self.progress(i, total, "正在写入 Word ...")

        outdir.mkdir(parents=True, exist_ok=True)
        safe = re.sub(r'[\\/:*?"<>|]', "_", dname)
        dp = outdir / f"{safe}_聊天记录.docx"
        doc.save(str(dp))
        self.progress(total, total, "Word 已生成")
        return dp

    def write_txt(self, msgs: list[dict], target: dict, me_wxid: str,
                  outdir: Path, my_nick: str = "我",
                  rename: dict[str, str] | None = None) -> Path:
        rename = rename or {}
        dname = target["name"]
        outdir.mkdir(parents=True, exist_ok=True)
        safe = re.sub(r'[\\/:*?"<>|]', "_", dname)
        tp = outdir / f"{safe}_聊天记录.txt"
        with open(tp, "w", encoding="utf-8") as f:
            f.write(f"与 {dname} 的聊天记录（{len(msgs)} 条）\n")
            if msgs:
                f.write(f"{datetime.datetime.fromtimestamp(msgs[0]['time']):%Y-%m-%d %H:%M}"
                        f" ~ {datetime.datetime.fromtimestamp(msgs[-1]['time']):%Y-%m-%d %H:%M}\n")
            f.write("=" * 60 + "\n\n")
            for m in msgs:
                dt = datetime.datetime.fromtimestamp(m["time"])
                if m["sender"] == me_wxid:
                    who = my_nick
                elif m["sender"] == target["wxid"]:
                    who = dname
                else:
                    who = m["sender"] or "系统"
                f.write(f"[{dt:%Y-%m-%d %H:%M:%S}] {who}："
                        f"{norm_content(m['content'], m['type'], rename)}\n")
        return tp

    def _find_attach_root(self) -> Path | None:
        if not self.db_dir:
            return None
        cand = self.db_dir.parent / "msg"
        return cand if cand.is_dir() else None

    # ---- 5. 图片 ----
    # 微信 4.x 真实布局（本机实测）：
    #   msg/attach/<32位会话哈希>/<YYYY-MM>/Img/*.dat     图片（AES 加密），*_t.dat 是缩略图
    # 注意：attach 下**没有**按月顶层目录，月份在哈希目录内部。
    def _session_hash(self, wxid: str) -> str:
        return hashlib.md5(wxid.encode()).hexdigest()

    # ---- 5.1 图片：V2 .dat 解密 ----
    # 微信 4.x 的 Img/*.dat 是 V2 加密：
    #   [6B 07 08 56 32 08 07][4B aes_size LE][4B xor_size LE][1B pad]
    #   + AES-128-ECB 段 + 明文段 + XOR 段
    # 密钥不读进程内存，用账号 uin 派生（同一套算法 Windows/macOS 通用）：
    #   xor_key = uin & 0xFF
    #   aes_key = MD5(str(uin) + wxid).hexdigest()[:16]
    # uin 不在磁盘上，用 xor_key 把枚举压到 2^24，再用图片 magic 校验命中。
    V2_MAGIC = b"\x07\x08\x56\x32\x08\x07"
    IMG_TIMEOUT = 60          # 单个 ffmpeg 转换上限（秒）

    def _img_cache_path(self) -> Path:
        base = os.environ.get("LOCALAPPDATA") or str(Path.home())
        return Path(base) / "微信导出工具" / "image_key.json"

    def _load_cached_image_key(self, wxid: str):
        """返回 (xor_key, aes_key, uin)。uin 可能为空（老版本缓存只存了前两项）。"""
        try:
            data = json.loads(self._img_cache_path().read_text(encoding="utf-8"))
            item = data.get(wxid) or {}
            if item.get("xor") is not None and item.get("aes"):
                uin = item.get("uin")
                return (int(item["xor"]), str(item["aes"]),
                        int(uin) if uin else None)
        except Exception:
            pass
        return None

    def _save_cached_image_key(self, wxid: str, xor_key: int, aes_key: str,
                               uin: int | None = None) -> None:
        try:
            p = self._img_cache_path()
            data = {}
            if p.exists():
                try:
                    data = json.loads(p.read_text(encoding="utf-8"))
                except Exception:
                    data = {}
            data[wxid] = {"xor": xor_key, "aes": aes_key,
                          **({"uin": uin} if uin else {})}
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                         encoding="utf-8")
        except Exception:
            pass

    @staticmethod
    def _img_magic_ok(block: bytes) -> bool:
        """严格判断解密后是不是图片。宽松判断会被随机字节撞出假阳性。"""
        return (block[:3] == b"\xff\xd8\xff"
                or block[:8] == b"\x89PNG\r\n\x1a\n"
                or block[:4] == b"GIF8"
                or (block[:4] == b"RIFF" and block[8:12] == b"WEBP")
                or block[:4] == b"wxgf")

    @staticmethod
    def _aes_block(aes_key: str, block: bytes) -> bytes:
        from Crypto.Cipher import AES
        return AES.new(aes_key.encode("ascii")[:16], AES.MODE_ECB).decrypt(block)

    def _v2_samples(self, attach_dir: Path, n: int = 4) -> list[bytes]:
        """取几张原图首块（AES 段前 16 字节）当密钥验证模板。"""
        out: list[bytes] = []
        if not attach_dir.is_dir():
            return out
        for p in attach_dir.rglob("Img/*.dat"):
            if p.name.endswith("_t.dat"):
                continue
            try:
                raw = p.open("rb").read(31)
            except OSError:
                continue
            if len(raw) < 31 or raw[:6] != self.V2_MAGIC:
                continue
            out.append(raw[15:31])
            if len(out) >= n:
                break
        return out

    def _vote_xor_key(self, attach_dir: Path, sample: int = 40) -> int | None:
        """缩略图几乎都是 JPG，末字节 = 0xD9 ^ xor_key，投票取多数。"""
        from collections import Counter
        votes: Counter = Counter()
        if not attach_dir.is_dir():
            return None
        for p in attach_dir.rglob("Img/*_t.dat"):
            try:
                with p.open("rb") as fp:
                    if fp.read(6) != self.V2_MAGIC:
                        continue
                    fp.seek(-1, 2)
                    votes[fp.read(1)[0] ^ 0xD9] += 1
            except OSError:
                continue
            if sum(votes.values()) >= sample:
                break
        if not votes:
            return None
        return votes.most_common(1)[0][0]

    def _account_wxid(self) -> str:
        """登录账号的 wxid（从 db_storage 的上级目录名取，去掉随机后缀）。

        图片密钥按**账号**派生，不是按会话；用会话 wxid 去枚举永远不命中。
        """
        if not self.db_dir:
            return ""
        name = self.db_dir.parent.name
        m = re.match(r"^(wxid_[^_]+)", name, re.IGNORECASE)
        return m.group(1) if m else name

    def derive_image_key(self, wxid: str, attach_dir: Path,
                         progress: Callable[[float, str], None] | None = None,
                         cancel: Callable[[], bool] | None = None,
                         time_budget: float = 180.0) -> tuple[int, str, int | None] | None:
        """拿本账号的 (xor_key, aes_key, uin)。成功会写缓存，下次秒开。

        uin 就是表情包密钥要用的 seed，顺手一起缓存，免得再枚举一遍。
        """
        if cancel and cancel():
            raise ExportCancelled("已取消")
        cached = self._load_cached_image_key(wxid)
        if cached and cached[2] is not None:
            blocks = self._v2_samples(attach_dir, 2)
            if all(self._img_magic_ok(self._aes_block(cached[1], b)) for b in blocks):
                self.log("图片密钥：用缓存的")
                self.image_uin = cached[2]
                return cached
            self.log("图片密钥：缓存对不上，重新推导")
        elif cached:
            self.log("图片密钥：老缓存缺 uin（表情包要用），重新推导一次")
        samples = self._v2_samples(attach_dir, 4)
        if not samples:
            self.log("图片密钥：没找到可验证的 V2 样本，跳过图片")
            return None
        xor_key = self._vote_xor_key(attach_dir)
        if xor_key is None:
            self.log("图片密钥：缩略图投票失败，跳过图片")
            return None
        self.log(f"图片密钥：xor={xor_key}，开始枚举 uin（约 1-2 分钟）")
        suf = wxid.encode()
        started = time.time()
        tried = 0
        for uin in range(xor_key, 1 << 32, 256):
            if cancel and cancel():
                raise ExportCancelled("已取消")
            tried += 1
            if tried % 250000 == 0:
                if time.time() - started > time_budget:
                    self.log("图片密钥：超时未命中，跳过图片")
                    return None
                if progress:
                    progress(min(95.0, tried * 100.0 / (1 << 24)),
                             f"正在推导图片密钥 {tried * 100 // (1 << 24)}%")
            key = hashlib.md5(str(uin).encode() + suf).hexdigest()[:16]
            aes = self._aes_block(key, samples[0])
            if not self._img_magic_ok(aes):
                continue
            if all(self._img_magic_ok(self._aes_block(key, b)) for b in samples[1:]):
                self.log(f"图片密钥：命中 uin={uin}，用时 {time.time()-started:.0f}s")
                self._save_cached_image_key(wxid, xor_key, key, uin)
                self.image_uin = uin
                return xor_key, key, uin
        self.log("图片密钥：枚举完没命中，跳过图片")
        return None

    @staticmethod
    def _strip_pkcs7(data: bytes) -> bytes:
        if not data:
            return data
        pad = data[-1]
        if 0 < pad <= 16 and data[-pad:] == bytes([pad]) * pad:
            return data[:-pad]
        return data

    def decrypt_dat(self, raw: bytes, xor_key: int, aes_key: str) -> bytes | None:
        """V2 .dat -> 原始字节（可能是 JPEG/PNG/GIF/WEBP，也可能是 wxgf 容器）。"""
        if len(raw) < 0x0F + 16 or raw[:6] != self.V2_MAGIC:
            return None
        aes_size, xor_size = struct.unpack("<II", raw[6:14])
        payload = raw[15:]
        # 微信按 PKCS7 补齐整个密文块：aes_size 是 16 的倍数时还要多读一块，
        # 否则会从明文段里多吃 16 字节（实测会把 PNG 的 chunk 链截断）。
        aligned = aes_size + (16 - aes_size % 16)
        if aligned > len(payload):
            return None
        try:
            plain = self._strip_pkcs7(self._aes_block(aes_key, payload[:aligned]))
        except Exception:
            return None
        rest = payload[aligned:]
        if xor_size and xor_size <= len(rest):
            raw_part = rest[:len(rest) - xor_size]
            tail = bytes(b ^ xor_key for b in rest[len(rest) - xor_size:])
            data = plain + raw_part + tail
        else:
            data = plain + rest
        return self._trim_to_image_end(data)

    @staticmethod
    def _trim_to_image_end(data: bytes) -> bytes:
        """裁掉尾部零星脏字节。

        实测：JPEG 约七成文件末尾会多出 24 字节（PNG 不受影响）。
        按「最后一个结束标记」截断后，实测 270/270 张都能被正常解码。
        """
        if data[:3] == b"\xff\xd8\xff":
            if data[-2:] != b"\xff\xd9":
                j = data.rfind(b"\xff\xd9")
                if j > 0:
                    return data[:j + 2]
        elif data[:8] == b"\x89PNG\r\n\x1a\n":
            if b"IEND" not in data[-16:]:
                j = data.rfind(b"IEND")
                if j > 0:
                    return data[:j + 8]
        elif data[:4] == b"GIF8" and data[-1:] != b"\x3b":
            j = data.rfind(b"\x3b")
            if j > 0:
                return data[:j + 1]
        return data

    @staticmethod
    def _find_image_magic(data: bytes, limit: int = 4096):
        """wxgf 里可能直接嵌了 JPEG/PNG。返回 (偏移, 扩展名) 或 None。"""
        end = min(len(data) - 12, limit)
        for i in range(4, end):
            if data[i] == 0xFF and data[i + 1] == 0xD8 and data[i + 2] == 0xFF:
                return i, ".jpg"
            if data[i:i + 4] == b"\x89PNG":
                return i, ".png"
        return None

    @staticmethod
    def find_ffmpeg() -> str | None:
        cands = []
        env = os.environ.get("WXE_FFMPEG")
        if env:
            cands.append(env)
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            cands.append(str(Path(meipass) / "ffmpeg.exe"))
        here = Path(sys.executable).parent
        cands += [str(here / "ffmpeg.exe"), str(Path(__file__).parent / "ffmpeg.exe")]
        for c in cands:
            if c and Path(c).is_file():
                return c
        return shutil.which("ffmpeg")

    def wxgf_to_jpg(self, data: bytes) -> bytes | None:
        """wxgf 容器里是 HEVC 单帧，用 ffmpeg 转 JPG。没有 ffmpeg 就返回 None。"""
        ff = self.find_ffmpeg()
        if not ff:
            return None
        payload = data[16:]
        if len(payload) < 200:
            payload = data[4:]
        tmpd = Path(tempfile.mkdtemp(prefix="wxe_img_"))
        src, dst = tmpd / "a.hevc", tmpd / "a.jpg"
        try:
            src.write_bytes(payload)
            # CREATE_NO_WINDOW：GUI 里并发调 ffmpeg 时，不隐藏会一闪一闪弹一堆黑窗
            subprocess.run([ff, "-v", "error", "-y", "-i", str(src),
                            "-frames:v", "1", "-q:v", "2", str(dst)],
                           capture_output=True, timeout=self.IMG_TIMEOUT,
                           creationflags=_NO_WINDOW, startupinfo=_NO_WINDOW_SI)
            if dst.is_file() and dst.stat().st_size > 512:
                return dst.read_bytes()
        except Exception:
            return None
        finally:
            shutil.rmtree(tmpd, ignore_errors=True)
        return None

    # ---- 5.2 图片内嵌：把 type=3 消息对应到本机 .dat 并解出图片 ----
    # 对应关系（本机实测 2026-11-21）：
    #   消息 XML 的 md5 = 解密后文件的 md5；hardlink.db 的 image_hardlink_info_v4
    #   正好记着 md5 -> 本地文件名（文件名本身 ≠ md5，别直接按名字找）。
    # 只有缩略图留在本机时，用「文件字节数 = cdnthumblength + 31」兜底匹配。
    THUMB_OVERHEAD = 31

    def _image_hardlink_map(self, dec: Path) -> dict[str, str]:
        p = dec / "hardlink" / "hardlink.db"
        out: dict[str, str] = {}
        if not p.exists():
            return out
        try:
            con = open_db(p)
        except sqlite3.Error:
            return out
        try:
            for md5v, fn in con.execute(
                    "SELECT md5, file_name FROM image_hardlink_info_v4"):
                if md5v and fn:
                    out[str(md5v).lower()] = str(fn)
        except sqlite3.Error:
            pass
        finally:
            con.close()
        return out

    def _chat_image_files(self, attach_root: Path, wxid: str):
        """返回 (原图 {stem: Path}, 缩略图 {stem: Path})。"""
        base = attach_root / "attach" / self._session_hash(wxid)
        full: dict[str, Path] = {}
        thumb: dict[str, Path] = {}
        if not base.is_dir():
            return full, thumb
        for p in base.rglob("*.dat"):
            if p.parent.name != "Img":
                continue
            if p.stem.endswith("_t"):
                thumb[p.stem[:-2]] = p
            else:
                full[p.stem] = p
        return full, thumb

    @staticmethod
    def _img_xml(content) -> str:
        if isinstance(content, bytes) and content[:4] == bytes.fromhex("28b52ffd"):
            return _zstd(content) or ""
        return content.decode("utf-8", "replace") if isinstance(content, bytes) else (content or "")

    @classmethod
    def _img_msg_md5(cls, content) -> tuple[str, str]:
        """消息里 md5（本地文件）、originsourcemd5（原图）两个值。"""
        s = cls._img_xml(content)
        a = re.search(r'\bmd5="([0-9a-fA-F]{32})"', s)
        b = re.search(r'originsourcemd5="([0-9a-fA-F]{32})"', s)
        return (a.group(1).lower() if a else "", b.group(1).lower() if b else "")

    @classmethod
    def _img_thumb_len(cls, content) -> int:
        m = re.search(r'cdnthumblength="(\d+)"', cls._img_xml(content))
        return int(m.group(1)) if m else 0

    def _decode_image_file(self, p: Path, xor_key: int, aes_key: str):
        try:
            raw = p.read_bytes()
        except OSError:
            return None, None
        data = self.decrypt_dat(raw, xor_key, aes_key)
        if not data:
            return None, None
        if data[:4] == b"wxgf":
            hit = self._find_image_magic(data)
            if hit:
                return data[hit[0]:], hit[1]
            jpg = self.wxgf_to_jpg(data)
            return (jpg, ".jpg") if jpg else (None, None)
        ext = (".jpg" if data[:3] == b"\xff\xd8\xff" else
               ".png" if data[:8] == b"\x89PNG\r\n\x1a\n" else
               ".gif" if data[:4] == b"GIF8" else
               ".webp" if data[:4] == b"RIFF" and data[8:12] == b"WEBP" else None)
        return (data, ext) if ext else (None, None)

    @staticmethod
    def _shrink_image(data: bytes, max_px: int = 1600) -> bytes | None:
        """压成 JPEG，控制 Word 体积；压不了就返回 None 用原图。"""
        try:
            import io
            from PIL import Image
            im = Image.open(io.BytesIO(data))
            im.load()
            if im.mode in ("RGBA", "LA", "P"):
                rgba = im.convert("RGBA")
                bg = Image.new("RGB", rgba.size, (255, 255, 255))
                bg.paste(rgba, mask=rgba.split()[-1])
                im = bg
            elif im.mode != "RGB":
                im = im.convert("RGB")
            w, h = im.size
            if max(w, h) > max_px:
                r = max_px / float(max(w, h))
                im = im.resize((max(1, int(w * r)), max(1, int(h * r))),
                               Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=85, optimize=True)
            return buf.getvalue()
        except Exception:
            return None

    def attach_inline_images(self, msgs: list[dict], target_wxid: str,
                             dec: Path | None = None,
                             progress: Callable[[float, str], None] | None = None,
                             cancel: Callable[[], bool] | None = None) -> dict:
        """给每条 type=3 消息挂上 m["img"] = (图片字节, 扩展名)，供 Word 内嵌。"""
        dec = dec or self.dec_dir
        stats = {"内嵌": 0, "缩略图": 0, "缺失": 0, "跳过": 0}
        imgs = [m for m in msgs if m.get("type") == 3]
        if not imgs:
            return stats
        root = self._find_attach_root()
        if not root or not dec:
            stats["缺失"] = len(imgs)
            return stats
        full, thumb = self._chat_image_files(root, target_wxid)
        if not full and not thumb:
            self.log("图片：这个会话在本机没有图片文件")
            stats["缺失"] = len(imgs)
            return stats
        account = self._account_wxid() or target_wxid
        key = self.derive_image_key(account, root / "attach",
                                    progress=progress, cancel=cancel)
        if not key:
            stats["缺失"] = len(imgs)
            return stats
        xor_key, aes_key, _uin = key
        hl = self._image_hardlink_map(dec)
        by_size: dict[int, list[Path]] = {}
        for p in thumb.values():
            try:
                by_size.setdefault(p.stat().st_size, []).append(p)
            except OSError:
                continue
        used: set[Path] = set()
        total = len(imgs)
        self.log(f"图片消息 {total} 条，本机原图 {len(full)} 张、缩略图 {len(thumb)} 张")
        for i, m in enumerate(imgs, 1):
            if cancel and cancel():
                raise ExportCancelled("已取消")
            if progress and (i % 5 == 0 or i == total):
                progress(i * 100.0 / total, f"正在解密图片 {i}/{total}")
            md5v, osmd5 = self._img_msg_md5(m.get("content"))
            path = None
            for v in (md5v, osmd5):
                fn = hl.get(v) if v else None
                if fn:
                    stem = fn[:-4] if fn.lower().endswith(".dat") else fn
                    hit = full.get(stem)
                    if hit is not None and hit not in used:
                        path = hit
                        break
            if path is None:
                n = self._img_thumb_len(m.get("content"))
                cands = [p for p in by_size.get(n + self.THUMB_OVERHEAD, [])
                         if p not in used] if n else []
                if cands:
                    t = m.get("time") or 0
                    cands.sort(key=lambda p: abs(p.stat().st_mtime - t))
                    path = cands[0]
            if path is None:
                stats["缺失"] += 1
                continue
            used.add(path)
            data, ext = self._decode_image_file(path, xor_key, aes_key)
            if not data:
                stats["跳过"] += 1
                continue
            if ext != ".gif" or len(data) > 300_000:
                shr = self._shrink_image(data)
                if shr:
                    data, ext = shr, ".jpg"
            m["img"] = (data, ext)
            stats["缩略图" if path.name.endswith("_t.dat") else "内嵌"] += 1
        self.log(f"  图片：内嵌 {stats['内嵌']} 张、用缩略图 {stats['缩略图']} 张、"
                 f"本机没有 {stats['缺失']} 张、解不开 {stats['跳过']} 张")
        return stats

    # ---- 表情消息（type=47）：直接内嵌到 Word ----
    STICKER_MAX_PX = 320

    @classmethod
    def _sticker_md5(cls, content) -> str:
        """表情消息 XML 里的 md5（= 本机表情文件的内容 md5）。"""
        s = cls._img_xml(content)
        m = re.search(r'<emoji\b[^>]*\bmd5="([0-9a-fA-F]{32})"', s)
        if not m:
            m = re.search(r'\bmd5="([0-9a-fA-F]{32})"', s)
        return m.group(1).lower() if m else ""

    def _emoticon_uin(self) -> int | None:
        """表情密钥的 seed（= 图片用的 uin）；图片没解过就读缓存的。"""
        if self.image_uin is not None:
            return self.image_uin
        acc = self._account_wxid()
        if not acc:
            return None
        c = self._load_cached_image_key(acc)
        return c[2] if c else None

    def _scan_sticker_blobs(self, want: set[str], key: bytes,
                            progress: Callable[[float, str], None] | None = None,
                            cancel: Callable[[], bool] | None = None) -> dict[str, bytes]:
        """在 business/emoticon 与 cache\\<月>\\Emoticon 里捞出这批 md5 的表情。

        只解密文件名/容器分片命中 want 的，不全量解密；缩略图只在没原图时兜底。
        """
        body: dict[str, bytes] = {}
        thumb: dict[str, bytes] = {}
        if not self.db_dir:
            return body
        acc_root = self.db_dir.parent
        emo_dir = acc_root / "business" / "emoticon"
        pat = re.compile(r"^([0-9a-f]{32})(?:\.thumb)?$")
        cands: list[tuple[Path, bool, bool]] = []   # (路径, 是否缩略, 是否容器)
        for sub in ("Persist", "PersistStore", "Thumb", "ThumbStore"):
            d = emo_dir / sub
            if not d.is_dir():
                continue
            for p in d.rglob("*"):
                m = pat.match(p.name)
                if p.is_file() and m and m.group(1) in want:
                    cands.append((p, sub.startswith("Thumb"), sub == "PersistStore"))
        cache = acc_root / "cache"
        if cache.is_dir():
            for p in cache.rglob("Emoticon/*/*"):
                m = pat.match(p.name)
                if p.is_file() and m and m.group(1) in want:
                    cands.append((p, False, False))
        total = len(cands) or 1
        for i, (p, is_thumb, is_container) in enumerate(cands, 1):
            if cancel and cancel():
                raise ExportCancelled("已取消")
            if progress and (i % 10 == 0 or i == total):
                progress(i * 100.0 / total, f"正在解密表情 {i}/{total}")
            data = self._decrypt_emo_file(p, key)
            if not data:
                continue
            tgt = thumb if is_thumb else body
            if is_container:
                for part in self._split_emo_container(data):
                    h = hashlib.md5(part).hexdigest()
                    if h in want:
                        tgt.setdefault(h, part)
            else:
                tgt.setdefault(pat.match(p.name).group(1), data)
        for k, v in thumb.items():
            body.setdefault(k, v)
        return body

    def _sticker_image(self, data: bytes):
        """表情文件 → (字节, 扩展名)；wxgf 转 JPG，大图/大 GIF 压成首帧 JPEG。"""
        fmt = self._emo_fmt(data)
        if fmt == "wxgf":
            jpg = self.wxgf_to_jpg(data)
            if not jpg:
                return None
            data, fmt = jpg, "jpg"
        if not fmt:
            return None
        if fmt == "gif" and len(data) <= 400_000:
            return data, ".gif"
        if fmt == "png" and len(data) <= 120_000:
            return data, ".png"
        # ffmpeg 出的 JPG 可能以 COM 段开头（ffd8fffe），Word/python-docx 不认，
        # 所以除了标准 JFIF/Exif 的一律交给 PIL 重编码成标准 JPG。
        if (fmt == "jpg" and len(data) <= 120_000
                and data[:4] in (b"\xff\xd8\xff\xe0", b"\xff\xd8\xff\xe1")):
            return data, ".jpg"
        shr = self._shrink_image(data, max_px=self.STICKER_MAX_PX)
        if shr:
            return shr, ".jpg"
        return data, "." + fmt

    def attach_inline_stickers(self, msgs: list[dict],
                               progress: Callable[[float, str], None] | None = None,
                               cancel: Callable[[], bool] | None = None) -> dict:
        """给每条 type=47 消息挂 m["img"] = (表情字节, 扩展名)，供 Word 内嵌。"""
        stats = {"内嵌": 0, "缺失": 0, "跳过": 0, "种类": 0, "总种类": 0}
        st = [m for m in msgs if m.get("type") == 47]
        if not st:
            return stats
        want: dict[str, int] = {}
        for m in st:
            v = self._sticker_md5(m.get("content"))
            if v:
                want[v] = want.get(v, 0) + 1
        stats["总种类"] = len(want)
        if not want:
            stats["缺失"] = len(st)
            self.log("表情：消息里没有 md5，跳过")
            return stats
        acc = self._account_wxid()
        uin = self._emoticon_uin()
        root = self._find_attach_root()
        if uin is None and acc and root:
            got = self.derive_image_key(acc, root / "attach",
                                        progress=progress, cancel=cancel)
            if got:
                self.image_uin = got[2]
                uin = got[2]
        if uin is None or not acc:
            self.log("表情：没拿到解密种子（需要微信里打开过图片/表情），这次用「[表情]」占位")
            stats["缺失"] = len(st)
            return stats
        key = self._emoticon_key(uin, acc)
        blobs = self._scan_sticker_blobs(set(want), key, progress, cancel)
        stats["种类"] = len(blobs)
        total = len(st)
        for i, m in enumerate(st, 1):
            if cancel and cancel():
                raise ExportCancelled("已取消")
            if progress and (i % 20 == 0 or i == total):
                progress(i * 100.0 / total, f"正在内嵌表情 {i}/{total}")
            v = self._sticker_md5(m.get("content"))
            data = blobs.get(v) if v else None
            if not data:
                stats["缺失"] += 1
                continue
            out = self._sticker_image(data)
            if not out:
                stats["跳过"] += 1
                continue
            m["img"] = out
            m["img_kind"] = "sticker"
            stats["内嵌"] += 1
        self.log(f"  表情：内嵌 {stats['内嵌']} 条（本机有 {stats['种类']}/{stats['总种类']} 种），"
                 f"本机没有 {stats['缺失']} 条、解不开 {stats['跳过']} 条")
        return stats

    # ---- 表情包（账号级，AES-128-CBC key=IV）----
    # 布局：business/emoticon/{Persist,PersistStore,Thumb,ThumbStore}，文件名是内容 md5。
    # 密钥：key = md5(seed + 账号wxid + "EMOTICON") 的原始前 16 字节，seed 就是图片用的 uin。
    # PersistStore 里是多个表情拼在一起的容器，按魔数 + 结构结束位置切开。
    EMO_SUBDIRS = ("Persist", "PersistStore")

    @staticmethod
    def _emo_fmt(dec: bytes) -> str:
        if dec[:8] == b"\x89PNG\r\n\x1a\n":
            return "png"
        if dec[:6] in (b"GIF87a", b"GIF89a"):
            return "gif"
        if dec[:3] == b"\xff\xd8\xff":
            return "jpg"
        if dec[:4] == b"wxgf":
            return "wxgf"
        if dec[:4] == b"RIFF":
            return "webp"
        return ""

    @staticmethod
    def _emoticon_key(uin: int, wxid: str) -> bytes:
        d = hashlib.md5(f"{uin}{wxid}EMOTICON".encode()).hexdigest()
        return bytes.fromhex(d)[:16]

    def _decrypt_emo_file(self, path: Path, key: bytes) -> bytes | None:
        from Crypto.Cipher import AES

        try:
            raw = path.read_bytes()
        except OSError:
            return None
        if len(raw) < 32 or len(raw) % 16:
            return None
        try:
            dec = AES.new(key, AES.MODE_CBC, key).decrypt(raw)
        except Exception:
            return None
        dec = self._strip_pkcs7(dec)
        return dec if self._emo_fmt(dec) else None

    # --- 容器分割：只认能完整解析到结构结束的魔数，避免压缩数据里的误匹配 ---
    @staticmethod
    def _emo_split_png(data: bytes, start: int) -> int | None:
        if data[start:start + 8] != b"\x89PNG\r\n\x1a\n":
            return None
        pos, n = start + 8, len(data)
        while pos + 8 <= n:
            ln = struct.unpack(">I", data[pos:pos + 4])[0]
            typ = data[pos + 4:pos + 8]
            if pos + 12 + ln > n:
                return None
            if typ == b"IEND":
                return pos + 12 + ln
            pos += 12 + ln
        return None

    @staticmethod
    def _emo_split_gif(data: bytes, start: int) -> int | None:
        if data[start:start + 6] not in (b"GIF87a", b"GIF89a"):
            return None
        pos, n = start + 6, len(data)
        if pos + 7 > n:
            return None
        flags = data[pos + 4]
        pos += 7
        if flags & 0x80:
            pos += 3 * (2 ** ((flags & 0x07) + 1))
        while pos < n:
            b = data[pos]
            if b == 0x3B:
                return pos + 1
            if b == 0x21:
                pos += 2
            elif b == 0x2C:
                if pos + 10 > n:
                    return None
                pos += 10
                if data[pos - 1] & 0x80:
                    pos += 3 * (2 ** ((data[pos - 1] & 0x07) + 1))
                if pos >= n:
                    return None
                pos += 1
            else:
                return None
            while pos < n:
                sz = data[pos]
                pos += 1
                if sz == 0:
                    break
                pos += sz
            if pos > n:
                return None
        return None

    @staticmethod
    def _emo_split_jpeg(data: bytes, start: int) -> int | None:
        if data[start:start + 2] != b"\xff\xd8":
            return None
        pos, n = start + 2, len(data)
        while pos + 2 <= n:
            if data[pos] != 0xFF:
                pos += 1
                continue
            marker = data[pos + 1]
            if marker == 0x00:
                pos += 2
                continue
            if marker == 0xD9:
                return pos + 2
            if marker == 0x01 or 0xD0 <= marker <= 0xD7:
                pos += 2
                continue
            if pos + 4 > n:
                return None
            seg = struct.unpack(">H", data[pos + 2:pos + 4])[0]
            if seg < 2 or pos + 2 + seg > n:
                return None
            pos += 2 + seg
        return None

    @classmethod
    def _emo_magic_at(cls, data: bytes, i: int) -> str:
        if data[i:i + 8] == b"\x89PNG\r\n\x1a\n":
            return "png"
        if data[i:i + 6] in (b"GIF87a", b"GIF89a"):
            return "gif"
        if data[i:i + 2] == b"\xff\xd8":
            return "jpg"
        if data[i:i + 4] == b"wxgf":
            return "wxgf"
        if data[i:i + 4] == b"RIFF":
            return "webp"
        return ""

    @classmethod
    def _split_emo_container(cls, data: bytes) -> list[bytes]:
        parts, i, n = [], 0, len(data)
        while i < n - 3:
            ext = cls._emo_magic_at(data, i)
            if not ext:
                i += 1
                continue
            if ext == "png":
                end = cls._emo_split_png(data, i)
            elif ext == "gif":
                end = cls._emo_split_gif(data, i)
            elif ext == "jpg":
                end = cls._emo_split_jpeg(data, i)
            else:  # wxgf / webp 结构不定，切到下一个魔数
                end = next((j for j in range(i + 4, n - 3)
                            if cls._emo_magic_at(data, j)), n)
            if end is None or end - i < 16:
                i += 1
                continue
            parts.append(data[i:end])
            i = end
        return parts

    def _emo_db_meta(self) -> dict:
        """从已解密的 emoticon.db 取包名/图注/顺序；拿不到就返回空（退回 md5 命名）。"""
        meta = {"packs": {}, "captions": {}, "files": {}, "fav": set()}
        db = (self.dec_dir / "emoticon" / "emoticon.db") if self.dec_dir else None
        if not db or not db.is_file():
            self.log("表情包：没拿到 emoticon.db，按 md5 命名")
            return meta
        try:
            con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
            cur = con.cursor()
            meta["packs"] = dict(cur.execute(
                "SELECT package_id_, package_name_ FROM kStoreEmoticonPackageTable"))
            for pid, md5, _lang, cap in cur.execute(
                    "SELECT package_id_, md5_, language_, caption_ "
                    "FROM kStoreEmoticonCaptionsTable"):
                meta["captions"][md5] = cap
            for pid, md5, _t, sort in cur.execute(
                    "SELECT package_id_, md5_, type_, sort_order_ "
                    "FROM kStoreEmoticonFilesTable"):
                meta["files"][md5] = (pid, sort)
            meta["fav"] = {r[0] for r in cur.execute(
                "SELECT md5 FROM kNonStoreEmoticonTable")}
            con.close()
        except Exception as e:
            self.log(f"表情包：读 emoticon.db 失败（{type(e).__name__}），按 md5 命名")
        return meta

    def gather_emoticons(self, outdir: Path,
                         cancel: Callable[[], bool] | None = None,
                         progress: Callable[[float, str], None] | None = None) -> dict:
        """把本账号的表情包解密导出到 表情包/，按包名 + 图注命名。"""
        stats = {"表情": 0, "跳过": 0}
        if not self.db_dir:
            return stats
        emo_dir = self.db_dir.parent / "business" / "emoticon"
        if not emo_dir.is_dir():
            self.log("表情包：本机没有 business/emoticon 目录，跳过")
            return stats
        uin = self.image_uin
        acc = self._account_wxid()
        if uin is None or not acc:
            self.log("表情包：没拿到 uin，跳过（图片密钥推导成功后才可导出）")
            return stats
        key = self._emoticon_key(uin, acc)

        pat = re.compile(r"^([0-9a-f]{32})(\.thumb)?$")
        targets: list[tuple[str, Path]] = []
        for sub in self.EMO_SUBDIRS:
            d = emo_dir / sub
            if not d.is_dir():
                continue
            for p in d.rglob("*"):
                m = pat.match(p.name)
                if p.is_file() and m:
                    targets.append((m.group(1), p))
        if not targets:
            self.log("表情包：没找到加密文件")
            return stats

        # 先验一把密钥：随便解一个，看魔数对不对
        probe = None
        for _md5, p in targets[:6]:
            probe = self._decrypt_emo_file(p, key)
            if probe:
                break
        if not probe:
            self.log("表情包：密钥对不上（微信版本可能不同），跳过")
            return stats

        files: dict[str, bytes] = {}
        total = len(targets)
        for i, (md5, p) in enumerate(targets, 1):
            if cancel and cancel():
                raise ExportCancelled("已取消")
            if progress and (i % 20 == 0 or i == total):
                progress(i * 100.0 / total, f"正在解密表情包 {i}/{total}")
            data = self._decrypt_emo_file(p, key)
            if not data:
                stats["跳过"] += 1
                continue
            if p.parent.name == "PersistStore":
                for part in self._split_emo_container(data):
                    files.setdefault(hashlib.md5(part).hexdigest(), part)
            else:
                files.setdefault(md5, data)

        meta = self._emo_db_meta()
        dst_root = outdir / "表情包"
        clean = lambda s: (re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", str(s)).strip()
                           or "未命名")
        for md5, data in files.items():
            if cancel and cancel():
                raise ExportCancelled("已取消")
            fmt = self._emo_fmt(data)
            if fmt == "wxgf":
                jpg = self.wxgf_to_jpg(data)
                if not jpg:
                    stats["跳过"] += 1
                    continue
                data, fmt = jpg, "jpg"
            if md5 in meta["files"]:
                pid, sort = meta["files"][md5]
                pkg = clean(meta["packs"].get(pid) or pid or "未命名包")
                cap = clean(meta["captions"].get(md5) or md5[:8])
                sub = dst_root / pkg
                name = f"{int(sort or 0):03d}_{cap}.{fmt}"
            elif md5 in meta["fav"]:
                sub = dst_root / "收藏"
                name = f"{md5[:8]}.{fmt}"
            else:
                sub = dst_root / "其他"
                name = f"{md5[:8]}.{fmt}"
            sub.mkdir(parents=True, exist_ok=True)
            try:
                (sub / name).write_bytes(data)
                stats["表情"] += 1
            except OSError:
                stats["跳过"] += 1
        self.log(f"  表情包：{stats['表情']} 个（{stats['跳过']} 个解不开）")
        return stats

    def cleanup(self):
        if self.tmp and self.tmp.exists():
            shutil.rmtree(self.tmp, ignore_errors=True)
            self.tmp = None