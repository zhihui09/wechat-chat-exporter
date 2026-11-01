# -*- coding: utf-8 -*-
"""在桌面目录生成 integrity.json 的临时副本（沙箱里写不了桌面，故先落 Temp）。"""
import hashlib
import json
from pathlib import Path

ROOT = Path(r"C:\Users\QQQQQ\Desktop\微信导出工具")
OUT = Path(r"C:\Users\QQQQQ\.dsh\tools\wx-export\integrity_new.json")

files = {}
for p in sorted(ROOT.rglob("*")):
    rel = p.relative_to(ROOT).as_posix()
    if not p.is_file() or p.name in ("integrity.json", "log.txt"):
        continue
    if ".bak-" in rel:          # 旧版本备份目录不参与校验
        continue
    files[rel] = hashlib.sha256(p.read_bytes()).hexdigest()

OUT.write_text(json.dumps({"sha256": files}, ensure_ascii=False), encoding="utf-8")
print("entries", len(files), "size", OUT.stat().st_size)
