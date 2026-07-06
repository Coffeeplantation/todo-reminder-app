Attribute VB_Name = "TodoReminder"
Option Explicit

' ============================================================
'  Todoリマインダー VBAモジュール
'  初回セットアップ: Alt+F8 -> SetupButtons を実行
' ============================================================

' ---- 初回セットアップ：シートにボタンを追加 ----
Sub SetupButtons()
    Dim ws As Worksheet
    Dim shp As Shape
    Set ws = ThisWorkbook.Sheets("Todo表")

    ' 既存ボタン削除
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = "btn_check" Or s.Name = "btn_reset" Then s.Delete
    Next s

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

    MsgBox "セットアップ完了！ボタンが追加されました。", vbInformation, "セットアップ完了"
End Sub

' ---- メイン：リマインダーチェック＆Outlook送信 ----
Sub CheckAndSendReminders()
    Dim wsTodo As Worksheet
    Dim wsContacts As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim today As Date
    Dim dueDate As Date
    Dim reminderDays As Long
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
    Set wsTodo = ThisWorkbook.Sheets("Todo表")
    Set wsContacts = ThisWorkbook.Sheets("名簿リスト")
    On Error GoTo 0

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
        If reminderDays <= 0 Then reminderDays = 3  ' デフォルト3日前

        ' 完了済み・送信済みはスキップ
        If statusVal = "完了" Then GoTo NextRow
        If sentFlag = "済" Then skipCount = skipCount + 1: GoTo NextRow

        ' 残り日数を計算
        daysLeft = DateDiff("d", today, dueDate)

        If daysLeft < 0 Then
            ' 期日超過
            wsTodo.Cells(i, 7).Value = "期日超過"
            wsTodo.Cells(i, 7).Interior.Color = RGB(255, 99, 71)
        ElseIf daysLeft <= reminderDays Then
            ' リマインド対象：メールアドレス検索
            emailAddr = GetEmailByName(wsContacts, assigneeName)

            If emailAddr <> "" Then
                If SendReminderEmail(emailAddr, assigneeName, taskName, dueDate, daysLeft) Then
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
        MsgBox msg, vbInformation, "送信完了"
    Else
        MsgBox "送信対象のタスクはありませんでした。" & vbCrLf & _
               "（スキップ済: " & skipCount & " 件）", vbInformation, "チェック完了"
    End If
    Exit Sub

SheetError:
    MsgBox "「Todo表」または「名簿リスト」シートが見つかりません。" & vbCrLf & _
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
Function SendReminderEmail(emailAddr As String, personName As String, _
                            taskName As String, dueDate As Date, daysLeft As Long) As Boolean
    Dim outlookApp As Object
    Dim mail As Object
    Dim subject As String
    Dim body As String

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
           "このメールはTodoリマインダーシステムより自動送信されています。"

    With mail
        .To = emailAddr
        .Subject = subject
        .Body = body
        .Send   ' 直接送信。確認画面を出したい場合は .Display に変更
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

    Set wsTodo = ThisWorkbook.Sheets("Todo表")
    lastRow = wsTodo.Cells(wsTodo.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        wsTodo.Cells(i, 7).Value = ""
        wsTodo.Cells(i, 7).Interior.ColorIndex = xlNone
    Next i

    MsgBox "リセット完了しました。", vbInformation, "リセット完了"
End Sub
