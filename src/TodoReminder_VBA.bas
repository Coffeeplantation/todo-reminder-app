Attribute VB_Name = "TodoReminder"
Option Explicit

' ============================================================
'  Todoリマインダー VBAモジュール
'  初回セットアップ: Alt+F8 -> SetupButtons を実行
'  シート構成: todoリスト（設定） / タスクリスト / 名簿
' ============================================================

Public gKanban As KanbanSheetCtrl

' ---- 初回セットアップ：シートにボタンを追加 ----
Sub SetupButtons()
    Dim ws As Worksheet
    Dim shp As Shape
    Set ws = ThisWorkbook.Sheets("todoリスト")

    ' 前回 Setup 時のシート保護が残っていても図形を削除・再生成できるように解除
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ' 既存ボタン・ブロック削除（廃止した btn_kanban も含む）
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = "btn_check" Or s.Name = "btn_reset" Or s.Name = "btn_kanban" _
           Or s.Name = "shp_settings_block" Then s.Delete
    Next s

    ' 設定項目（B4:C8）を角丸四角形の枠でひとまとまりに
    Dim rngBlock As Range
    Set rngBlock = ws.Range("B4:C8")

    Dim shpBlock As Shape
    Set shpBlock = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        rngBlock.Left - 6, rngBlock.Top - 6, rngBlock.Width + 12, rngBlock.Height + 12)
    shpBlock.Name = "shp_settings_block"
    shpBlock.Fill.Visible = msoFalse
    shpBlock.Line.ForeColor.RGB = RGB(43, 63, 99)
    shpBlock.Line.Weight = 1.75
    shpBlock.Shadow.Visible = False
    shpBlock.TextFrame.Characters().Text = ""
    shpBlock.Placement = xlMove
    shpBlock.ZOrder msoSendToBack

    ' チェックボタン（青）
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 680, 5, 155, 24)
    shp.Name = "btn_check"
    shp.TextFrame.Characters().Text = "リマインド送信チェック"
    shp.TextFrame.Characters().Font.Bold = True
    shp.TextFrame.Characters().Font.Size = 9
    shp.TextFrame.Characters().Font.Color = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(68, 114, 196)
    shp.Line.Visible = False
    shp.OnAction = "CheckAndSendReminders"

    ' リセットボタン（緑）
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 680, 35, 155, 20)
    shp.Name = "btn_reset"
    shp.TextFrame.Characters().Text = "送信済みフラグをリセット"
    shp.TextFrame.Characters().Font.Size = 8
    shp.TextFrame.Characters().Font.Color = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(112, 173, 71)
    shp.Line.Visible = False
    shp.OnAction = "ResetSentFlags"

    ' Kanban ボード（シート内埋め込み）をセットアップ
    SetupKanbanSheet

    MsgBox "セットアップ完了！ボタンと Kanban ボードが追加されました。", vbInformation, "セットアップ完了"
End Sub

' ---- Kanban ボード（todoリスト シート内埋め込み）をセットアップ／再構築 ----
Sub SetupKanbanSheet()
    ' 再実行のたびに必ず新しいインスタンスを作る
    ' （クラスコードを更新した後に古いインスタンスが残ると状態が混在するため）
    Set gKanban = New KanbanSheetCtrl
    gKanban.Setup ThisWorkbook.Sheets("todoリスト")
End Sub

' ---- メイン：リマインダーチェック＆Outlook送信 ----
Sub CheckAndSendReminders()
    Dim wsMain As Worksheet
    Dim wsTodo As Worksheet
    Dim wsContacts As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim today As Date
    Dim dueDate As Date
    Dim reminderDays As Long
    Dim defaultReminderDays As Long
    Dim autoMailOn As Boolean
    Dim signature As String
    Dim assigneeName As String
    Dim taskName As String
    Dim statusVal As String
    Dim sentFlag As String
    Dim emailAddr As String
    Dim daysLeft As Long
    Dim sentCount As Integer
    Dim skipCount As Integer

    today = Date
    sentCount = 0
    skipCount = 0

    On Error GoTo SheetError
    Set wsMain = ThisWorkbook.Sheets("todoリスト")
    Set wsTodo = ThisWorkbook.Sheets("タスクリスト")
    Set wsContacts = ThisWorkbook.Sheets("名簿")
    On Error GoTo 0

    ' ---- todoリスト（設定シート）の設定を読み込み ----
    autoMailOn = (Trim(wsMain.Range("C8").Value) <> "OFF")
    defaultReminderDays = wsMain.Range("C4").Value
    If defaultReminderDays <= 0 Then defaultReminderDays = 3
    signature = wsMain.Range("C5").Value
    If signature = "" Then signature = "Todoリマインダーシステム"

    lastRow = wsTodo.Cells(wsTodo.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Todoデータがありません。", vbInformation, "チェック完了"
        Exit Sub
    End If

    For i = 2 To lastRow
        ' --- データ取得 ---
        If wsTodo.Cells(i, 1).Value = "" Then Exit For  ' No.が空なら終了

        taskName     = wsTodo.Cells(i, 2).Value   ' B: タスク名
        assigneeName = wsTodo.Cells(i, 3).Value   ' C: 担当者名
        statusVal    = wsTodo.Cells(i, 6).Value   ' F: ステータス
        sentFlag     = wsTodo.Cells(i, 7).Value   ' G: 送信済み

        ' 期日が空ならスキップ
        If wsTodo.Cells(i, 4).Value = "" Then GoTo NextRow

        On Error Resume Next
        dueDate = CDate(wsTodo.Cells(i, 4).Value)   ' D: 期日
        If Err.Number <> 0 Then Err.Clear: GoTo NextRow
        On Error GoTo 0

        reminderDays = wsTodo.Cells(i, 5).Value     ' E: リマインド日数前
        If reminderDays <= 0 Then reminderDays = defaultReminderDays  ' todoリストの既定値

        ' 完了済み・送信済みはスキップ
        If statusVal = "完了" Then GoTo NextRow
        If sentFlag = "済" Then skipCount = skipCount + 1: GoTo NextRow

        ' 残り日数を計算
        daysLeft = DateDiff("d", today, dueDate)

        If daysLeft < 0 Then
            ' 期日超過（自動メール設定に関わらず判定・表示する）
            wsTodo.Cells(i, 7).Value = "期日超過"
            wsTodo.Cells(i, 7).Interior.Color = RGB(255, 99, 71)
        ElseIf daysLeft <= reminderDays Then
            ' リマインド対象：メールアドレス検索
            emailAddr = GetEmailByName(wsContacts, assigneeName)

            If emailAddr <> "" Then
                If SendReminderEmail(autoMailOn, emailAddr, assigneeName, taskName, dueDate, daysLeft, signature) Then
                    wsTodo.Cells(i, 7).Value = "済"
                    wsTodo.Cells(i, 7).Interior.Color = RGB(144, 238, 144)
                    sentCount = sentCount + 1
                End If
            Else
                wsTodo.Cells(i, 7).Value = "アドレス不明"
                wsTodo.Cells(i, 7).Interior.Color = RGB(255, 165, 0)
            End If
        End If

NextRow:
    Next i

    Dim msg As String
    If sentCount > 0 Then
        msg = sentCount & " 件のリマインドメールを送信しました。"
        If skipCount > 0 Then msg = msg & vbCrLf & "（送信済みスキップ: " & skipCount & " 件）"
        If Not autoMailOn Then msg = msg & vbCrLf & "※自動メール送信はOFFのため、実際のメールは送信されていません（表示のみ更新）。"
    Else
        msg = "送信対象のタスクはありませんでした。" & vbCrLf & _
              "（スキップ済: " & skipCount & " 件）"
    End If
    MsgBox msg, vbInformation, "チェック完了"
    Exit Sub

SheetError:
    MsgBox "「todoリスト」「タスクリスト」「名簿」のいずれかのシートが見つかりません。" & vbCrLf & _
           "シート名を確認してください。", vbCritical, "シートエラー"
End Sub

' ---- 名前でメールアドレスを検索 ----
Function GetEmailByName(ws As Worksheet, personName As String) As String
    Dim lastRow As Long
    Dim i As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) = Trim(personName) Then
            GetEmailByName = Trim(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    GetEmailByName = ""
End Function

' ---- Outlook メール送信 ----
Function SendReminderEmail(autoMailOn As Boolean, emailAddr As String, personName As String, _
                            taskName As String, dueDate As Date, daysLeft As Long, _
                            signature As String) As Boolean
    Dim outlookApp As Object
    Dim mail As Object
    Dim subject As String
    Dim body As String

    If Not autoMailOn Then
        ' 自動メール送信OFF：実際の送信は行わず、判定・表示のみ更新する
        SendReminderEmail = True
        Exit Function
    End If

    On Error GoTo SendError

    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)  ' olMailItem

    subject = "【リマインド】" & taskName & _
              "（期日：" & Format(dueDate, "yyyy/mm/dd") & "）"

    If daysLeft = 0 Then
        body = personName & " 様" & vbCrLf & vbCrLf & _
               "本日が期日のタスクがあります。ご確認をお願いいたします。" & vbCrLf & vbCrLf
    Else
        body = personName & " 様" & vbCrLf & vbCrLf & _
               "期日まで残り " & daysLeft & " 日となりました。ご対応をお願いいたします。" & vbCrLf & vbCrLf
    End If

    body = body & _
           "━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & _
           "  タスク名：" & taskName & vbCrLf & _
           "  期　　日：" & Format(dueDate, "yyyy年mm月dd日") & vbCrLf & _
           "  残り日数：" & daysLeft & " 日" & vbCrLf & _
           "━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & _
           "このメールは" & signature & "より自動送信されています。"

    With mail
        .To = emailAddr
        .Subject = subject
        .Body = body
        .Display   ' テスト用：ドラフト表示（本番運用時は .Send に戻す）
    End With

    SendReminderEmail = True

    Set mail = Nothing
    Set outlookApp = Nothing
    Exit Function

SendError:
    MsgBox "メール送信エラー" & vbCrLf & _
           "担当者: " & personName & vbCrLf & _
           "タスク: " & taskName & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "送信エラー"
    SendReminderEmail = False
End Function

' ---- 「送信済み」フラグを一括リセット ----
Sub ResetSentFlags()
    Dim wsTodo As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim answer As VbMsgBoxResult

    answer = MsgBox("送信済みフラグを全てリセットしますか？" & vbCrLf & _
                    "（「済」「アドレス不明」「期日超過」を全て空白に戻します）", _
                    vbYesNo + vbQuestion, "フラグリセット確認")

    If answer = vbNo Then Exit Sub

    Set wsTodo = ThisWorkbook.Sheets("タスクリスト")
    lastRow = wsTodo.Cells(wsTodo.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        wsTodo.Cells(i, 7).Value = ""
        wsTodo.Cells(i, 7).Interior.ColorIndex = xlNone
    Next i

    MsgBox "リセット完了しました。", vbInformation, "リセット完了"
End Sub
