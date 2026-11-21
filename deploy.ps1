$ErrorActionPreference = "Stop"
$src  = "C:\Users\QQQQQ\.dsh\tools\wx-export\dist\微信导出工具"
$dst  = "C:\Users\QQQQQ\Desktop\微信导出工具"
$py   = "C:\Users\QQQQQ\AppData\Local\Programs\Python\Python313\python.exe"

if (Test-Path "$dst\微信导出工具.exe") {
    Copy-Item "$dst\微信导出工具.exe" "$dst\微信导出工具.exe.bak-20261121" -Force
}
# _internal 整包换掉，避免上一版残留模块；桌面里用户自己的文件夹不动
if (Test-Path "$dst\_internal") { Remove-Item "$dst\_internal" -Recurse -Force }
Copy-Item "$src\_internal" $dst -Recurse -Force
Copy-Item "$src\微信导出工具.exe" "$dst\微信导出工具.exe" -Force
foreach ($f in @("使用说明.txt", "README.md")) {
    if (Test-Path "C:\Users\QQQQQ\.dsh\tools\wx-export\$f") {
        Copy-Item "C:\Users\QQQQQ\.dsh\tools\wx-export\$f" "$dst\$f" -Force
    }
}

# 重新生成 integrity.json（只收程序自身文件）
& $py -X utf8 -c "import sys; sys.path.insert(0, r'C:\Users\QQQQQ\.dsh\tools\wx-export'); import wxexport_core as c; from pathlib import Path; p=c.write_integrity_manifest(Path(r'$dst')); print('manifest', p)"
if ($LASTEXITCODE -ne 0) { throw "manifest failed" }

# 去掉低完整性标签，否则双击秒退
icacls "$dst" /setintegritylevel Medium /T /C /Q | Out-Null
"DEPLOY_OK"
