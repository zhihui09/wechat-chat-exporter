$ErrorActionPreference = "Stop"
$src  = "C:\Users\QQQQQ\.dsh\tools\wx-export\dist\微信导出工具"
$dst  = "C:\Users\QQQQQ\Desktop\微信导出工具"
$py   = "C:\Users\QQQQQ\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"

if (Test-Path "$dst\微信导出工具.exe") {
    Copy-Item "$dst\微信导出工具.exe" "$dst\微信导出工具.exe.bak-20261031b" -Force
}
# 直接覆盖（_internal 内容已核对一致，不再整体删除，避免锁文件权限问题）
Copy-Item "$src\*" $dst -Recurse -Force

# 重新生成 integrity.json
& $py -X utf8 -c "import sys; sys.path.insert(0, r'C:\Users\QQQQQ\.dsh\tools\wx-export'); import wxexport_core as c; from pathlib import Path; p=c.write_integrity_manifest(Path(r'$dst')); print('manifest', p)"
if ($LASTEXITCODE -ne 0) { throw "manifest failed" }

# 去掉低完整性标签，否则双击秒退
icacls "$dst" /setintegritylevel Medium /T /C /Q | Out-Null
"DEPLOY_OK"
