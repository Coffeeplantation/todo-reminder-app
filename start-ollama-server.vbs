' ローカルAI（Ollama）サーバーをバックグラウンドで起動する
' スタートアップ（HKCU Run キー "OllamaServe"）から wscript.exe 経由で呼ばれる。
' トレイアプリ（ollama app.exe）がサーバーを起動しないことがあるため、serve を直接起動する。
' CORS 許可（OLLAMA_ORIGINS=*）はユーザー環境変数に設定済みで、起動時に自動で引き継がれる。
Option Explicit

Dim sh, exe
Set sh = CreateObject("WScript.Shell")
exe = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Ollama\ollama.exe"

' ウィンドウ非表示(0)・待たない(False)。すでに起動済みならポート使用中で即終了するだけで無害
sh.Run """" & exe & """ serve", 0, False
