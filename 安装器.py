# -*- coding: utf-8 -*-
"""微信导出工具 · 安装器

打包成单文件 exe，别人下载双击即可安装：
  · 解压程序本体到 %LOCALAPPDATA%\\微信导出工具（不需要管理员）
  · 建桌面 + 开始菜单快捷方式
  · 生成卸载脚本
命令行（供自测用）：
  安装器.exe --silent <目录>      静默安装到指定目录
  安装器.exe --uninstall <目录>   静默卸载
"""
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile
from pathlib import Path

APP_NAME = "微信导出工具"
EXE_NAME = "微信导出工具.exe"

C_BG = "#F7F6F2"
C_CARD = "#FFFFFF"
C_TEXT = "#2C2F36"
C_SUB = "#8A9099"
C_ACCENT = "#07C160"
C_ACCENT_D = "#06A050"
FONT = "Microsoft YaHei UI"


# ---------------------------------------------------------------- 路径工具
def _meipass() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).parent))


def payload_path() -> Path:
    return _meipass() / "payload.zip"


def default_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / APP_NAME


def desktop_dir() -> Path:
    try:
        import winreg
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as k:
            v, _ = winreg.QueryValueEx(k, "Desktop")
            p = Path(os.path.expandvars(v))
            if p.is_dir():
                return p
    except Exception:
        pass
    p = Path.home() / "Desktop"
    return p if p.is_dir() else Path.home()


def start_menu_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / APP_NAME


# ---------------------------------------------------------------- 安装动作
def extract_payload(target: Path, progress=lambda pct, msg: None) -> None:
    target.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(payload_path()) as z:
        items = z.infolist()
        total = len(items) or 1
        for i, info in enumerate(items, 1):
            z.extract(info, target)
            if i % 40 == 0 or i == total:
                progress(int(i * 100 / total), f"正在解压 {i}/{total}")


def make_shortcut(lnk: Path, target_exe: Path, workdir: Path, icon: Path,
                  desc: str = "") -> None:
    """用 PowerShell 建 .lnk（避免任何第三方依赖）。"""
    lnk.parent.mkdir(parents=True, exist_ok=True)
    ps = f"""$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{lnk}')
$s.TargetPath='{target_exe}'
$s.WorkingDirectory='{workdir}'
$s.IconLocation='{icon}'
$s.Description='{desc}'
$s.Save()
"""
    f = Path(tempfile.mkdtemp()) / "mk.ps1"
    f.write_text("﻿" + ps, encoding="utf-8")
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                    "-File", str(f)], capture_output=True, timeout=60)


def write_uninstaller(target: Path) -> None:
    exe = target / EXE_NAME
    bat = target / "卸载.cmd"
    bat.write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        f'taskkill /f /im "{EXE_NAME}" >nul 2>&1\r\n'
        f'del /f /q "%USERPROFILE%\\Desktop\\{APP_NAME}.lnk" >nul 2>&1\r\n'
        f'rmdir /s /q "%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\{APP_NAME}"'
        " >nul 2>&1\r\n"
        f'cd /d "%TEMP%"\r\n'
        f'rmdir /s /q "{target}"\r\n'
        "echo 已卸载，窗口 3 秒后自动关闭...\r\n"
        "timeout /t 3 >nul\r\n",
        encoding="utf-8")


def install(target: Path, progress=lambda pct, msg: None) -> None:
    progress(1, "正在解压程序文件 ...")
    extract_payload(target, progress)
    exe = target / EXE_NAME
    if not exe.exists():
        raise RuntimeError("解压后没找到主程序，安装包可能不完整。")
    icon = target / "_internal" / "app.ico"

    if not os.environ.get("WXE_NO_SHORTCUT"):
        progress(96, "正在创建快捷方式 ...")
        make_shortcut(desktop_dir() / f"{APP_NAME}.lnk", exe, target, icon, APP_NAME)
        make_shortcut(start_menu_dir() / f"{APP_NAME}.lnk", exe, target, icon, APP_NAME)
    write_uninstaller(target)
    progress(100, "安装完成")


def uninstall(target: Path) -> None:
    for lnk in (desktop_dir() / f"{APP_NAME}.lnk", start_menu_dir() / f"{APP_NAME}.lnk"):
        try:
            lnk.unlink()
        except Exception:
            pass
    try:
        shutil.rmtree(start_menu_dir(), ignore_errors=True)
    except Exception:
        pass
    shutil.rmtree(target, ignore_errors=True)


# ---------------------------------------------------------------- 静默模式
def _say(msg: str) -> None:
    """打包成窗口程序后 sys.stdout 是 None，直接 print 会崩；统一走这里。"""
    try:
        if sys.stdout:
            sys.stdout.write(str(msg) + "\n")
    except Exception:
        pass
    try:
        with open(Path(tempfile.gettempdir()) / "wxsetup_out.txt", "a",
                  encoding="utf-8") as f:
            f.write(str(msg) + "\n")
    except Exception:
        pass


def silent(args: list[str]) -> int:
    if not args:
        _say("用法: --silent <目录> | --uninstall <目录>")
        return 2
    target = Path(args[1]) if len(args) > 1 else default_dir()
    _say(f"start {args[0]} {target}")
    if args[0] == "--uninstall":
        uninstall(target)
        _say("UNINSTALL_OK")
        return 0
    try:
        install(target, lambda p, m: _say(f"{p}% {m}"))
    except Exception as e:
        import traceback
        _say("INSTALL_FAIL " + repr(e))
        _say(traceback.format_exc())
        return 1
    ok = (target / EXE_NAME).exists()
    _say(("INSTALL_OK " if ok else "INSTALL_FAIL missing exe ") + str(target))
    return 0 if ok else 1


# ---------------------------------------------------------------- 图形界面
def gui() -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(f"{APP_NAME} 安装程序")
    root.configure(bg=C_BG)
    root.geometry("520x360")
    root.resizable(False, False)
    try:
        ico = _meipass() / "app.ico"
        if ico.exists():
            root.iconbitmap(default=str(ico))
    except Exception:
        pass

    tk.Label(root, text=f"安装 {APP_NAME}", bg=C_BG, fg=C_TEXT,
             font=(FONT, 16, "bold")).pack(anchor="w", padx=28, pady=(24, 4))
    tk.Label(root, text="导出微信一对一聊天记录为 Word / TXT（需要电脑版微信 4.x）",
             bg=C_BG, fg=C_SUB, font=(FONT, 9)).pack(anchor="w", padx=28)

    box = tk.Frame(root, bg=C_CARD, highlightthickness=1,
                   highlightbackground="#E6E3DC")
    box.pack(fill="x", padx=28, pady=18)
    tk.Label(box, text="安装位置", bg=C_CARD, fg=C_SUB,
             font=(FONT, 9)).pack(anchor="w", padx=14, pady=(12, 2))
    row = tk.Frame(box, bg=C_CARD)
    row.pack(fill="x", padx=14, pady=(0, 14))
    var = tk.StringVar(value=str(default_dir()))
    ent = tk.Entry(row, textvariable=var, relief="flat", bg="#F4F3EF",
                   fg=C_TEXT, font=(FONT, 9), insertbackground=C_TEXT)
    ent.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

    def browse():
        d = filedialog.askdirectory(initialdir=str(default_dir()))
        if d:
            var.set(str(Path(d) / APP_NAME) if Path(d).name != APP_NAME else d)

    tk.Button(row, text="浏览", command=browse, relief="flat", bg="#EDEBE5",
              fg=C_TEXT, font=(FONT, 9), bd=0, padx=12, pady=5,
              activebackground="#E2DFD8").pack(side="left")

    pb = ttk.Progressbar(root, mode="determinate", maximum=100)
    pb.pack(fill="x", padx=28)
    status = tk.Label(root, text="点击「开始安装」创建快捷方式并复制文件",
                      bg=C_BG, fg=C_SUB, font=(FONT, 9))
    status.pack(anchor="w", padx=28, pady=(6, 0))

    btn_row = tk.Frame(root, bg=C_BG)
    btn_row.pack(fill="x", padx=28, pady=16, side="bottom")

    st = {"busy": False}

    def on_install():
        if st["busy"]:
            return
        st["busy"] = True
        btn.configure(state="disabled", text="安装中 ...")

        def prog(p, m):
            pb["value"] = p
            status.configure(text=m)

        def work():
            try:
                install(Path(var.get()), prog)
                root.after(0, done, True, "")
            except Exception as e:
                root.after(0, done, False, str(e))

        threading.Thread(target=work, daemon=True).start()

    def done(ok, err):
        st["busy"] = False
        btn.configure(state="normal", text="重新安装")
        if not ok:
            status.configure(text="安装失败")
            messagebox.showerror(APP_NAME, f"安装失败：\n{err}")
            return
        pb["value"] = 100
        status.configure(text="安装完成 ✓")
        if messagebox.askyesno(
                APP_NAME,
                "安装完成！\n\n桌面和开始菜单都有快捷方式了。\n\n现在打开吗？\n\n"
                "提示：第一次运行会弹一次系统的「是否允许更改」，"
                "那是因为读微信数据需要管理员权限，点「是」即可。"):
            try:
                subprocess.Popen([str(Path(var.get()) / EXE_NAME)],
                                 cwd=str(Path(var.get())))
            except Exception as e:
                messagebox.showerror(APP_NAME, f"启动失败：{e}")

    btn = tk.Button(btn_row, text="开始安装", command=on_install, relief="flat",
                    bg=C_ACCENT, fg="white", font=(FONT, 11, "bold"), bd=0,
                    activebackground=C_ACCENT_D, activeforeground="white",
                    padx=26, pady=8, cursor="hand2")
    btn.pack(side="right")
    tk.Button(btn_row, text="取消", command=root.destroy, relief="flat",
              bg="#EDEBE5", fg=C_TEXT, font=(FONT, 9), bd=0, padx=18, pady=8,
              activebackground="#E2DFD8").pack(side="right", padx=(0, 10))

    root.mainloop()
    return 0


if __name__ == "__main__":
    argv = sys.argv[1:]
    if argv and argv[0] in ("--silent", "--uninstall"):
        sys.exit(silent(argv))
    sys.exit(gui())
