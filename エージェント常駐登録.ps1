# エージェント常駐登録.ps1 — ログオン時にTodoリマインダー常駐エージェントを自動起動する
# スタートアップフォルダにショートカットを作成します（管理者権限は不要）。
# 解除するには: スタートアップフォルダ（shell:startup）の「Todoリマインダー常駐エージェント」を削除
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbs = Join-Path $scriptDir "エージェントを開始.vbs"
$startup = [Environment]::GetFolderPath("Startup")
$lnkPath = Join-Path $startup "Todoリマインダー常駐エージェント.lnk"

$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($lnkPath)
$lnk.TargetPath = "wscript.exe"
$lnk.Arguments = "`"$vbs`""
$lnk.Description = "Todoリマインダー常駐エージェント（ログオン時自動起動）"
$lnk.Save()

Write-Host "✅ ログオン時の自動起動を登録しました: $lnkPath"
Write-Host "今すぐ起動します...（すでに起動中なら2重起動はされません）"
Start-Process wscript.exe -ArgumentList "`"$vbs`""
