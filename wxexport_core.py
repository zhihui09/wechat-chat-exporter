#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""微信聊天记录导出 - 核心库

把整个流程封装成可被 GUI 调用的类：
  1. 扫描微信数据库目录、从进程内存提取密钥、解密数据库
  2. 列出联系人与群聊
  3. 导出指定会话为 Word / TXT
  4. 从本地存储归集该会话的图片、视频、文件
"""
from __future__ import annotations

import atexit
import datetime
import functools
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
import zlib
import traceback
from pathlib import Path
from typing import Callable, Iterable

VENDOR = Path(__file__).resolve().parent / "vendor" / "wechat4-db-decrypt"
PAGE_SZ = 4096


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
    """给打包目录生成 integrity.json（除自身外全部文件的 sha256）。"""
    files = {}
    for p in sorted(root.rglob("*")):
        # log.txt 是运行时日志，每次启动都会改，不能进清单
        if p.is_file() and p.name not in ("integrity.json", "log.txt"):
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
    def write_word(self, msgs: list[dict], target: dict, me_wxid: str,
                   outdir: Path, my_nick: str = "我",
                   rename: dict[str, str] | None = None,
                   cancel: Callable[[], bool] | None = None) -> Path:
        from docx import Document
        from docx.shared import Pt, RGBColor

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
            bd = p.add_run(norm_content(m["content"], m["type"], rename))
            bd.font.size = Pt(10)
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

    # ---- 5. 媒体文件归集 ----
    # 微信 4.x 真实布局（本机实测）：
    #   msg/attach/<32位会话哈希>/<YYYY-MM>/Img/*.dat     图片（AES 加密）
    #   msg/attach/<32位会话哈希>/<YYYY-MM>/Video/*       视频
    #   msg/attach/<32位会话哈希>/<YYYY-MM>/Thumb/*       缩略图
    #   msg/file/<YYYY-MM>/<原文件名>                     收到的文件（明文）
    #   msg/video/<YYYY-MM>/*                             视频
    # 注意：attach 下**没有**按月顶层目录，月份在哈希目录内部。
    MEDIA_KINDS = {
        "图片": ("attach", None, ("Img",), (".dat",)),
        "视频": ("video", None, None, (".mp4", ".mov", ".mkv", ".avi")),
        "文件": ("file", None, None, None),
    }

    def _session_hash(self, wxid: str) -> str:
        return hashlib.md5(wxid.encode()).hexdigest()

    def gather_media(self, msgs: list[dict], outdir: Path,
                     target_wxid: str = "", limit: int = 5000,
                     cancel: Callable[[], bool] | None = None) -> dict:
        """把该会话的图片/视频/文件复制到 outdir/media，并按类型分好。

        只按「会话哈希 + 月份」定位，不做逐条消息对齐；不改动任何原文件。
        """
        root = self._find_attach_root()
        stats = {"图片": 0, "视频": 0, "文件": 0, "语音": 0}
        if not root:
            self.log("找不到附件目录（msg/），跳过媒体归集")
            return stats

        stamps = {int(m["time"]) for m in msgs if m.get("time")}
        months = {datetime.datetime.fromtimestamp(s).strftime("%Y-%m")
                  for s in stamps}
        if not months:
            self.log("消息里没有时间戳，跳过媒体归集")
            return stats
        self.log(f"涉及 {len(months)} 个月份，在 {root} 下归集附件")

        want = {k: outdir / "media" / k for k in ("视频", "文件")}
        for d in want.values():
            d.mkdir(parents=True, exist_ok=True)

        # 图片（msg/attach/<hash>/<month>/Img/*.dat）暂不处理：
        # 4.x 的 .dat 是 wxgf 封装 + AES，需要另一套密钥，归集出来也打不开。

        # ---- 视频 / 文件：按月顶层目录 ----
        for kind, subdir, _unused, exts in (
                ("视频", "video", None, (".mp4", ".mov", ".mkv", ".avi")),
                ("文件", "file", None, None)):
            base = root / subdir
            if not base.is_dir():
                continue
            for month in months:
                md = base / month
                if not md.is_dir():
                    continue
                for f in md.rglob("*"):
                    if cancel and cancel():
                        raise ExportCancelled("已取消")
                    if stats[kind] >= limit:
                        break
                    if not f.is_file():
                        continue
                    if exts and f.suffix.lower() not in exts:
                        continue
                    dst = want[kind] / f.name
                    if dst.exists():
                        dst = want[kind] / f"{f.parent.name}_{f.name}"
                    try:
                        shutil.copyfile(f, dst)   # copyfile 不继承只读属性
                        stats[kind] += 1
                    except (OSError, PermissionError):
                        continue
            self.log(f"  {kind}：{stats[kind]} 个")

        return stats

    def cleanup(self):
        if self.tmp and self.tmp.exists():
            shutil.rmtree(self.tmp, ignore_errors=True)
            self.tmp = None