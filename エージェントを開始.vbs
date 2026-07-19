' エージェントを開始.vbs — Todoリマインダー常駐エージェントを黒い窓なしで起動する
' ダブルクリックで起動。すでに起動中なら何もしない（ポート8787が使用中のため2重起動しない）。
' 停止するにはタスクマネージャーで node.exe を終了するか、PCを再起動してください。
Option Explicit

Dim shell, fso, scriptDir
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' 0 = ウィンドウ非表示, False = 待たずに戻る
shell.Run "cmd /c node """ & scriptDir & "\agent\agent.js"" >> """ & scriptDir & "\agent\data\agent-stdout.log"" 2>&1", 0, False

WScript.Sleep 1500
' ブラウザでアプリを開く（エージェントが配信するURL。連携された状態で使える）
shell.Run "http://127.0.0.1:8787/", 1, False
