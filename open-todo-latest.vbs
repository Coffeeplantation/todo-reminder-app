' Todoリマインダー（ブラウザ版）を最新版に更新してから開くランチャー
' デスクトップのショートカットから wscript.exe 経由で呼ばれる。
' 1. リポジトリを git pull で最新化（早送りできない場合やオフライン時はスキップして続行）
' 2. Chrome で browser/index.html を開く
Option Explicit

Dim sh, repo, chrome, page
Set sh = CreateObject("WScript.Shell")

repo = "C:\Users\katsu\Documents\todo-reminder-app"
chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
page = repo & "\browser\index.html"

' ウィンドウ非表示(0)で pull 完了を待ってから(True)開く
sh.Run "cmd /c cd /d """ & repo & """ && git pull --ff-only", 0, True

sh.Run """" & chrome & """ """ & page & """", 1, False
