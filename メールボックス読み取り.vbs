' メールボックス読み取り（Todoリマインダー用・完全ローカル）
' ---------------------------------------------------------------
' 指定したOutlookのメールフォルダ（メールボックス）を読み、メールの内容を
' テキストファイル「メールボックス書き出し.txt」としてデスクトップに保存します。
' 保存したファイルを Todoリマインダーの「📥 自動タスク作成」へドラッグ＆ドロップすると、
' ローカルAIがメールから予定を抽出し、現行予定表（タスクボード）へ追加できます。
'
' ・データはPCの中だけで処理されます（ネットへは一切送信しません）
' ・クラシック版Outlook（デスクトップアプリ）が必要です
' 使い方：このファイルをダブルクリック → フォルダ名と日数を入力
' ---------------------------------------------------------------
Option Explicit

Dim folderName, daysBack, inputDays
folderName = InputBox( _
  "読み取るメールフォルダ名を入力してください。" & vbCrLf & vbCrLf & _
  "・そのままOKで「受信トレイ」を読み取ります" & vbCrLf & _
  "・受信トレイ内のサブフォルダは名前だけでも探します（例：プロジェクトA）", _
  "メールボックス読み取り", "受信トレイ")
If folderName = "" Then WScript.Quit ' キャンセル

inputDays = InputBox( _
  "何日前までのメールを読み取りますか？（半角数字）", _
  "メールボックス読み取り", "30")
If inputDays = "" Then WScript.Quit
If Not IsNumeric(inputDays) Then
  MsgBox "日数は半角数字で入力してください。", vbExclamation, "メールボックス読み取り"
  WScript.Quit
End If
daysBack = CLng(inputDays)

Dim outlookApp, ns, targetFolder
On Error Resume Next
Set outlookApp = CreateObject("Outlook.Application")
If Err.Number <> 0 Then
  MsgBox "Outlookを起動できませんでした。" & vbCrLf & _
    "クラシック版Outlook（デスクトップアプリ）がインストールされているか確認してください。" & vbCrLf & _
    "（「新しいOutlook」のみの環境では利用できません）", vbCritical, "メールボックス読み取り"
  WScript.Quit
End If
On Error GoTo 0

Set ns = outlookApp.GetNamespace("MAPI")

' フォルダを探す：既定の受信トレイ → 受信トレイ配下 → 全ストアを再帰検索
Dim inbox
Set inbox = ns.GetDefaultFolder(6) ' 6 = olFolderInbox
If folderName = "受信トレイ" Or LCase(folderName) = "inbox" Then
  Set targetFolder = inbox
Else
  Set targetFolder = FindFolder(inbox, folderName)
  If targetFolder Is Nothing Then
    Dim store
    For Each store In ns.Folders
      Set targetFolder = FindFolder(store, folderName)
      If Not targetFolder Is Nothing Then Exit For
    Next
  End If
End If

If targetFolder Is Nothing Then
  MsgBox "フォルダ「" & folderName & "」が見つかりませんでした。" & vbCrLf & _
    "Outlookのフォルダ一覧に表示されている名前をそのまま入力してください。", vbExclamation, "メールボックス読み取り"
  WScript.Quit
End If

' 指定日数内のメールを新しい順に読み取り
Dim items, item, sinceDate, count, maxCount, text
sinceDate = DateAdd("d", -daysBack, Now)
Set items = targetFolder.Items
items.Sort "[ReceivedTime]", True ' 新しい順
maxCount = 200 ' 読み取り上限（AI処理の負荷保護）
count = 0
text = ""

For Each item In items
  If count >= maxCount Then Exit For
  On Error Resume Next
  If item.Class = 43 Then ' 43 = olMail
    If item.ReceivedTime >= sinceDate Then
      text = text & "=== MAIL ===" & vbCrLf
      text = text & "件名: " & item.Subject & vbCrLf
      text = text & "差出人: " & item.SenderName & vbCrLf
      text = text & "受信日時: " & FormatDateTime(item.ReceivedTime, 0) & vbCrLf
      text = text & "本文:" & vbCrLf
      text = text & TrimBody(item.Body) & vbCrLf & vbCrLf
      count = count + 1
    Else
      ' 新しい順に並んでいるため、期間より古いメールが出たら打ち切り
      Exit For
    End If
  End If
  On Error GoTo 0
Next

If count = 0 Then
  MsgBox "フォルダ「" & targetFolder.Name & "」に、直近 " & daysBack & " 日以内のメールはありませんでした。", vbInformation, "メールボックス読み取り"
  WScript.Quit
End If

' UTF-8（BOMなし）でデスクトップへ保存
Dim shell, outPath
Set shell = CreateObject("WScript.Shell")
outPath = shell.SpecialFolders("Desktop") & "\メールボックス書き出し.txt"
SaveUtf8 outPath, text

MsgBox "フォルダ「" & targetFolder.Name & "」から " & count & " 件のメールを書き出しました。" & vbCrLf & vbCrLf & _
  "保存先：" & outPath & vbCrLf & vbCrLf & _
  "Todoリマインダーの「📥 自動タスク作成」にこのファイルをドラッグ＆ドロップすると、" & vbCrLf & _
  "ローカルAIがメールから予定を読み取ります。", vbInformation, "メールボックス読み取り"

' ---- 補助関数 ----

' フォルダ名（完全一致・大文字小文字無視）を再帰的に探す
Function FindFolder(parent, name)
  Dim f, found
  Set FindFolder = Nothing
  For Each f In parent.Folders
    If LCase(f.Name) = LCase(name) Then
      Set FindFolder = f
      Exit Function
    End If
    Set found = FindFolder(f, name)
    If Not found Is Nothing Then
      Set FindFolder = found
      Exit Function
    End If
  Next
End Function

' 本文の整形：引用・署名以降を含め長すぎる場合は切り詰める（1通あたり最大4000文字）
Function TrimBody(body)
  Dim s
  s = Replace(body, vbCr & vbLf, vbLf)
  s = Replace(s, vbCr, vbLf)
  s = Replace(s, vbLf, vbCrLf)
  If Len(s) > 4000 Then s = Left(s, 4000) & vbCrLf & "（以下省略）"
  TrimBody = s
End Function

' UTF-8（BOMなし）でテキストを保存
Sub SaveUtf8(path, content)
  Dim streamUtf8, streamNoBom
  Set streamUtf8 = CreateObject("ADODB.Stream")
  streamUtf8.Type = 2 ' テキスト
  streamUtf8.Charset = "utf-8"
  streamUtf8.Open
  streamUtf8.WriteText content
  ' BOM（先頭3バイト）を取り除いてバイナリ保存し直す
  streamUtf8.Position = 0
  streamUtf8.Type = 1 ' バイナリ
  streamUtf8.Position = 3
  Set streamNoBom = CreateObject("ADODB.Stream")
  streamNoBom.Type = 1
  streamNoBom.Open
  streamUtf8.CopyTo streamNoBom
  streamNoBom.SaveToFile path, 2 ' 上書き
  streamNoBom.Close
  streamUtf8.Close
End Sub
