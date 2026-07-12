' Todoリマインダー（デスクトップアプリ版）を最新版に更新してから開くランチャー
' デスクトップのショートカットから wscript.exe 経由で呼ばれる。
' 1. リポジトリを git pull で最新化（早送りできない場合やオフライン時はスキップして続行）
' 2. Chrome のアプリモード（ブラウザ枠なしの専用ウィンドウ）で browser/index.html を開く
Option Explicit

Dim sh, repo, chrome, pageUrl
Set sh = CreateObject("WScript.Shell")

repo = "C:\Users\katsu\Documents\todo-reminder-app"
chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
pageUrl = "file:///" & Replace(repo, "\", "/") & "/browser/index.html"

' ウィンドウ非表示(0)で pull 完了を待ってから(True)開く。オフライン時は失敗してもそのまま続行
sh.Run "cmd /c cd /d """ & repo & """ && git pull --ff-only", 0, True

' --app= でアプリモード起動（タブやアドレスバーのない独立ウィンドウ）
sh.Run """" & chrome & """ --app=""" & pageUrl & """", 1, False
