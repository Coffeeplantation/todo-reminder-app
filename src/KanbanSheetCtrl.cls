VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "KanbanSheetCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  todoリスト シート上に埋め込む Kanban ボード（ActiveX 版）
'  frmKanban（廃止）の機能をシート内に移植したもの
' ============================================================

Private m_Ws            As Worksheet
Private m_Helpers(80)   As Object
Private m_HCount        As Integer
Private m_Lists(2)      As MSForms.ListBox
Private m_Headers(2)    As Shape
Private m_SelCol        As Integer
Private m_SelIdx        As Integer

Private m_Pills(30)     As MSForms.Label
Private m_PillCount     As Integer
Private m_DragAssignee  As String
Private lblGhost        As MSForms.Label

Private m_DragColIdx    As Integer   ' -1 = not dragging a card
Private m_DragItemIdx   As Integer
Private m_DragArmed     As Boolean
Private m_DragVisible   As Boolean
Private m_DragStartX    As Single
Private m_DragStartY    As Single

Private m_BaseLeft      As Single
Private m_BaseTop       As Single
Private m_Scale         As Single       ' フルボードの拡大率（1 = 等倍）

Private m_BoardNames    As Collection   ' フル Kanban ボードを構成する図形名
Private m_SummaryNames  As Collection   ' 簡易表示（一覧）を構成する図形名
Private m_SummaryList   As MSForms.ListBox

' ---- フルボードの基準レイアウト定数（scale = 1 のときの値） ----
Private Const BASE_TITLEBAR_Y As Single = -60
Private Const BASE_TITLEBAR_H As Single = 24
Private Const BASE_HINT_Y As Single = -30
Private Const BASE_COL_W As Single = 262
Private Const BASE_COL_GAP As Single = 274
Private Const BASE_HDR_H As Single = 26
Private Const BASE_LIST_Y As Single = 30
Private Const BASE_LIST_H As Single = 300
Private Const BASE_BTNADD_Y As Single = 334
Private Const BASE_BTNADD_H As Single = 24
Private Const BASE_ARROW_W As Single = 20
Private Const BASE_ARROW_H As Single = 22
Private Const BASE_BOTTOM_Y As Single = 364
Private Const BASE_BOTTOM_H As Single = 26
Private Const BASE_RESIZE_SIZE As Single = 16

' ---- 移動／リサイズのドラッグ状態 ----
Private m_MoveArmed     As Boolean
Private m_MoveStartX    As Single
Private m_MoveStartY    As Single
Private m_MoveOrigLeft  As Single
Private m_MoveOrigTop   As Single

Private m_ResizeArmed   As Boolean
Private m_ResizeStartX  As Single
Private m_ResizeOrigScale As Single

Public Sub Setup(ws As Worksheet)
    Set m_Ws = ws
    m_SelCol = -1: m_SelIdx = -1: m_HCount = 0
    m_DragColIdx = -1
    m_PillCount = 0
    m_Scale = 1
    m_MoveArmed = False: m_ResizeArmed = False
    Set m_BoardNames = New Collection
    Set m_SummaryNames = New Collection

    ' 前回 Setup 時にシート保護をかけている場合、図形の削除・再生成ができるように解除
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ClearExisting

    ' 基準位置：設定ブロック（B4:C8）の下
    m_BaseLeft = ws.Range("B11").Left
    m_BaseTop = ws.Range("B11").Top

    Dim titles(2) As String
    titles(0) = "  未完了"
    titles(1) = "  進行中"
    titles(2) = "  完了"

    Dim bgCol(2) As Long
    bgCol(0) = RGB(68, 114, 196)
    bgCol(1) = RGB(220, 120, 0)
    bgCol(2) = RGB(60, 150, 60)

    Dim i As Integer
    For i = 0 To 2
        Dim cx As Single: cx = m_BaseLeft + i * 274

        Dim shpH As Shape
        Set shpH = ws.Shapes.AddShape(msoShapeRectangle, cx, m_BaseTop, 262, 26)
        shpH.Name = "kb_hdr" & i
        shpH.TextFrame.Characters().Text = titles(i)
        shpH.TextFrame.Characters().Font.Bold = True
        shpH.TextFrame.Characters().Font.Size = 10
        shpH.TextFrame.Characters().Font.Color = RGB(255, 255, 255)
        shpH.Fill.ForeColor.RGB = bgCol(i)
        shpH.Line.Visible = False
        Set m_Headers(i) = shpH
        m_BoardNames.Add shpH.Name

        Dim oleLb As OLEObject
        Set oleLb = ws.OLEObjects.Add(ClassType:="Forms.ListBox.1", _
            Left:=cx, Top:=m_BaseTop + 30, Width:=262, Height:=300)
        oleLb.Name = "kb_list" & i
        Dim lb As MSForms.ListBox
        Set lb = oleLb.Object
        lb.Font.Size = 9
        lb.BackColor = RGB(252, 254, 255)
        lb.ColumnCount = 5
        lb.ColumnWidths = "110 pt;75 pt;52 pt;0 pt;0 pt"
        Set m_Lists(i) = lb
        m_BoardNames.Add oleLb.Name

        Dim oleAdd As OLEObject
        Set oleAdd = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
            Left:=cx, Top:=m_BaseTop + 334, Width:=262, Height:=24)
        oleAdd.Name = "kb_btnA" & i
        Dim bAdd As MSForms.CommandButton
        Set bAdd = oleAdd.Object
        bAdd.Caption = "+ タスク追加"
        bAdd.BackColor = bgCol(i)
        bAdd.ForeColor = RGB(255, 255, 255)
        bAdd.Font.Size = 9
        WireBtn bAdd
        m_BoardNames.Add oleAdd.Name

        Dim hLb As New KanbanHelper
        Set hLb.lb = lb
        hLb.colIdx = i
        Set hLb.formRef = Me
        Set m_Helpers(m_HCount) = hLb
        m_HCount = m_HCount + 1
    Next i

    AddArrow "→", m_BaseLeft + 255, m_BaseTop + 150, "kb_btnR0"
    AddArrow "←", m_BaseLeft + 255, m_BaseTop + 178, "kb_btnL0"
    AddArrow "→", m_BaseLeft + 529, m_BaseTop + 150, "kb_btnR1"
    AddArrow "←", m_BaseLeft + 529, m_BaseTop + 178, "kb_btnL1"

    Dim oleDel As OLEObject
    Set oleDel = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
        Left:=m_BaseLeft, Top:=m_BaseTop + 364, Width:=120, Height:=26)
    oleDel.Name = "kb_btnDel"
    Dim bDel As MSForms.CommandButton
    Set bDel = oleDel.Object
    bDel.Caption = "選択を削除"
    bDel.BackColor = RGB(180, 0, 0)
    bDel.ForeColor = RGB(255, 255, 255)
    bDel.Font.Size = 9
    WireBtn bDel
    m_BoardNames.Add oleDel.Name

    Dim oleClose As OLEObject
    Set oleClose = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
        Left:=m_BaseLeft + 130, Top:=m_BaseTop + 364, Width:=150, Height:=26)
    oleClose.Name = "kb_btnClose"
    Dim bClose As MSForms.CommandButton
    Set bClose = oleClose.Object
    bClose.Caption = "簡易表示に戻る"
    bClose.Font.Size = 9
    WireBtn bClose
    m_BoardNames.Add oleClose.Name

    Dim oleGhost As OLEObject
    Set oleGhost = ws.OLEObjects.Add(ClassType:="Forms.Label.1", _
        Left:=m_BaseLeft, Top:=m_BaseTop, Width:=10, Height:=24)
    oleGhost.Name = "kb_ghost"
    Set lblGhost = oleGhost.Object
    lblGhost.Visible = False
    lblGhost.Font.Bold = True
    lblGhost.Font.Size = 9
    lblGhost.TextAlign = fmTextAlignCenter
    lblGhost.SpecialEffect = fmSpecialEffectRaised
    m_BoardNames.Add oleGhost.Name

    ' ---- タイトルバー（ドラッグで移動） ----
    Dim oleTitle As OLEObject
    Set oleTitle = ws.OLEObjects.Add(ClassType:="Forms.Label.1", _
        Left:=m_BaseLeft, Top:=m_BaseTop + BASE_TITLEBAR_Y, _
        Width:=BoardWidth(1), Height:=BASE_TITLEBAR_H)
    oleTitle.Name = "kb_titlebar"
    Dim lblTitle As MSForms.Label
    Set lblTitle = oleTitle.Object
    lblTitle.Caption = "  ▤ Kanban ボード（ドラッグで移動 / 右下の角でリサイズ）"
    lblTitle.BackColor = RGB(43, 63, 99)
    lblTitle.ForeColor = RGB(255, 255, 255)
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 9
    lblTitle.TextAlign = fmTextAlignLeft
    lblTitle.MousePointer = fmMousePointerSizeAll
    m_BoardNames.Add oleTitle.Name

    Dim hTitle As New KanbanHelper
    Set hTitle.moveLbl = lblTitle
    Set hTitle.formRef = Me
    Set m_Helpers(m_HCount) = hTitle
    m_HCount = m_HCount + 1

    ' ---- リサイズハンドル（右下の角） ----
    Dim oleResize As OLEObject
    Set oleResize = ws.OLEObjects.Add(ClassType:="Forms.Label.1", _
        Left:=m_BaseLeft + BoardWidth(1) - BASE_RESIZE_SIZE, _
        Top:=m_BaseTop + BASE_BOTTOM_Y + BASE_BOTTOM_H, _
        Width:=BASE_RESIZE_SIZE, Height:=BASE_RESIZE_SIZE)
    oleResize.Name = "kb_resize"
    Dim lblResize As MSForms.Label
    Set lblResize = oleResize.Object
    lblResize.Caption = "◢"
    lblResize.BackColor = RGB(150, 150, 150)
    lblResize.ForeColor = RGB(255, 255, 255)
    lblResize.Font.Size = 8
    lblResize.TextAlign = fmTextAlignCenter
    lblResize.MousePointer = fmMousePointerSizeNWSE
    m_BoardNames.Add oleResize.Name

    Dim hResize As New KanbanHelper
    Set hResize.resizeLbl = lblResize
    Set hResize.formRef = Me
    Set m_Helpers(m_HCount) = hResize
    m_HCount = m_HCount + 1

    ' ---- ドラッグ中の枠プレビュー（通常は非表示） ----
    Dim shpOutline As Shape
    Set shpOutline = ws.Shapes.AddShape(msoShapeRectangle, m_BaseLeft, m_BaseTop, 10, 10)
    shpOutline.Name = "kb_outline"
    shpOutline.Fill.Visible = msoFalse
    shpOutline.Line.Visible = msoTrue
    shpOutline.Line.ForeColor.RGB = RGB(255, 140, 0)
    shpOutline.Line.Weight = 2
    shpOutline.Line.DashStyle = msoLineDash
    shpOutline.Visible = False

    BuildAssigneePills
    BuildSummaryList
    LoadFromExcel

    ' 設定ブロック（既定のリマインド日数前 等）とリマインド関連の2ボタンも
    ' フルボードと一緒に開閉する対象に含める
    m_BoardNames.Add "shp_settings_block"
    m_BoardNames.Add "btn_check"
    m_BoardNames.Add "btn_reset"

    ' 行の表示・非表示切替（設定ブロックの開閉）でセルに連動して図形が
    ' ズレたりサイズが変わったりしないよう、セルに依存しない配置に固定する
    FreezePlacement m_BoardNames
    FreezePlacement m_SummaryNames

    ToggleVisibility m_BoardNames, False
    ToggleVisibility m_SummaryNames, True

    ' 設定ブロックの行を非表示にすると B11 の位置がずれる可能性があるため、
    ' 非表示にした後に基準位置を測り直してから簡易表示を再配置する
    ToggleSettingsRows False
    m_BaseTop = m_Ws.Range("B11").Top
    LayoutSummary
    LoadSummary

    ProtectShapes
End Sub

' ユーザーがマウスで図形をドラッグして動かし、レイアウトを分解できないようにする。
' DrawingObjects:=True で図形の選択・ドラッグ移動・削除をUI操作から禁止する一方、
' Contents:=False（セル・コントロールの値は保護しない）＆ UserInterfaceOnly:=True
' （マクロからの変更は保護の影響を受けない）なので、
' ・ボタンのクリックやリストのクリック／選択
' ・本コード自身が行う再配置（SetRect 等）やドラッグ機能（タイトルバー移動、
'   リサイズ、担当者ピルのドラッグ、カードのドラッグ）
' はこれまで通り動作する。UserInterfaceOnly はブック保存時にリセットされる仕様のため、
' Setup（＝ Workbook_Open のたびに実行）で毎回かけ直す。
Private Sub ProtectShapes()
    On Error Resume Next
    m_Ws.Protect DrawingObjects:=True, Contents:=False, Scenarios:=True, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub FreezePlacement(names As Collection)
    Dim nm As Variant
    For Each nm In names
        On Error Resume Next
        m_Ws.Shapes(CStr(nm)).Placement = xlFreeFloating
        On Error GoTo 0
    Next nm
End Sub

Private Sub ToggleVisibility(names As Collection, vis As Boolean)
    Dim nm As Variant
    For Each nm In names
        On Error Resume Next
        m_Ws.Shapes(CStr(nm)).Visible = vis
        On Error GoTo 0
    Next nm
End Sub

Private Sub ToggleSettingsRows(vis As Boolean)
    m_Ws.Rows("4:8").Hidden = Not vis
End Sub

Private Function BoardWidth(s As Single) As Single
    BoardWidth = 2 * BASE_COL_GAP * s + BASE_COL_W * s
End Function

Private Function BoardHeightTotal(s As Single) As Single
    BoardHeightTotal = (-BASE_TITLEBAR_Y + BASE_BOTTOM_Y + BASE_BOTTOM_H) * s
End Function

Private Sub SetRect(nm As String, l As Single, t As Single, w As Single, h As Single)
    On Error Resume Next
    Dim sh As Shape
    Set sh = m_Ws.Shapes(nm)
    sh.Left = l
    sh.Top = t
    sh.Width = w
    sh.Height = h
    On Error GoTo 0
End Sub

' 現在の m_BaseLeft / m_BaseTop / m_Scale に基づき、フルボードの全図形を再配置
Private Sub LayoutBoard()
    Dim s As Single: s = m_Scale
    Dim i As Integer
    For i = 0 To 2
        Dim cx As Single: cx = m_BaseLeft + i * BASE_COL_GAP * s
        SetRect "kb_hdr" & i, cx, m_BaseTop, BASE_COL_W * s, BASE_HDR_H * s
        SetRect "kb_list" & i, cx, m_BaseTop + BASE_LIST_Y * s, BASE_COL_W * s, BASE_LIST_H * s
        SetRect "kb_btnA" & i, cx, m_BaseTop + BASE_BTNADD_Y * s, BASE_COL_W * s, BASE_BTNADD_H * s
    Next i

    SetRect "kb_btnR0", m_BaseLeft + 255 * s, m_BaseTop + 150 * s, BASE_ARROW_W * s, BASE_ARROW_H * s
    SetRect "kb_btnL0", m_BaseLeft + 255 * s, m_BaseTop + 178 * s, BASE_ARROW_W * s, BASE_ARROW_H * s
    SetRect "kb_btnR1", m_BaseLeft + 529 * s, m_BaseTop + 150 * s, BASE_ARROW_W * s, BASE_ARROW_H * s
    SetRect "kb_btnL1", m_BaseLeft + 529 * s, m_BaseTop + 178 * s, BASE_ARROW_W * s, BASE_ARROW_H * s

    SetRect "kb_btnDel", m_BaseLeft, m_BaseTop + BASE_BOTTOM_Y * s, 120 * s, BASE_BOTTOM_H * s
    SetRect "kb_btnClose", m_BaseLeft + 130 * s, m_BaseTop + BASE_BOTTOM_Y * s, 150 * s, BASE_BOTTOM_H * s

    SetRect "kb_hint", m_BaseLeft, m_BaseTop + BASE_HINT_Y * s, 230 * s, 24 * s
    SetRect "kb_titlebar", m_BaseLeft, m_BaseTop + BASE_TITLEBAR_Y * s, BoardWidth(s), BASE_TITLEBAR_H * s
    SetRect "kb_resize", m_BaseLeft + BoardWidth(s) - BASE_RESIZE_SIZE * s, _
        m_BaseTop + (BASE_BOTTOM_Y + BASE_BOTTOM_H) * s, BASE_RESIZE_SIZE * s, BASE_RESIZE_SIZE * s

    Dim px As Single: px = m_BaseLeft + 250 * s
    For i = 0 To m_PillCount - 1
        Dim pillW As Single: pillW = m_Pills(i).Width
        SetRect "kb_pill" & i, px, m_BaseTop + BASE_HINT_Y * s, pillW, 24 * s
        px = px + pillW + 8 * s
    Next i

    RescaleFonts
End Sub

Private Sub RescaleFonts()
    Dim fsz As Single: fsz = 9 * m_Scale
    If fsz < 6 Then fsz = 6
    If fsz > 16 Then fsz = 16
    Dim i As Integer
    For i = 0 To 2
        m_Lists(i).Font.Size = fsz
        m_Headers(i).TextFrame.Characters().Font.Size = 10 * m_Scale
    Next i
End Sub

Private Sub ShowOutline()
    m_Ws.Shapes("kb_outline").Visible = True
End Sub

Private Sub HideOutline()
    m_Ws.Shapes("kb_outline").Visible = False
End Sub

Public Sub OnMoveMouseDown(x As Single, y As Single)
    m_MoveArmed = True
    Dim tb As Shape: Set tb = m_Ws.Shapes("kb_titlebar")
    m_MoveStartX = tb.Left + x
    m_MoveStartY = tb.Top + y
    m_MoveOrigLeft = m_BaseLeft
    m_MoveOrigTop = m_BaseTop
    SetRect "kb_outline", m_BaseLeft, m_BaseTop + BASE_TITLEBAR_Y * m_Scale, BoardWidth(m_Scale), BoardHeightTotal(m_Scale)
    ShowOutline
End Sub

Public Sub OnMoveMouseMove(x As Single, y As Single)
    If Not m_MoveArmed Then Exit Sub
    Dim tb As Shape: Set tb = m_Ws.Shapes("kb_titlebar")
    Dim curX As Single, curY As Single
    curX = tb.Left + x: curY = tb.Top + y
    Dim dx As Single, dy As Single
    dx = curX - m_MoveStartX: dy = curY - m_MoveStartY
    Dim newLeft As Single, newTop As Single
    newLeft = m_MoveOrigLeft + dx
    newTop = m_MoveOrigTop + dy
    SetRect "kb_outline", newLeft, newTop + BASE_TITLEBAR_Y * m_Scale, BoardWidth(m_Scale), BoardHeightTotal(m_Scale)
End Sub

Public Sub OnMoveMouseUp(x As Single, y As Single)
    If Not m_MoveArmed Then Exit Sub
    m_MoveArmed = False
    Dim tb As Shape: Set tb = m_Ws.Shapes("kb_titlebar")
    Dim curX As Single, curY As Single
    curX = tb.Left + x: curY = tb.Top + y
    Dim dx As Single, dy As Single
    dx = curX - m_MoveStartX: dy = curY - m_MoveStartY
    m_BaseLeft = m_MoveOrigLeft + dx
    m_BaseTop = m_MoveOrigTop + dy
    HideOutline
    LayoutBoard
End Sub

Public Sub OnResizeMouseDown(x As Single, y As Single)
    m_ResizeArmed = True
    Dim rz As Shape: Set rz = m_Ws.Shapes("kb_resize")
    m_ResizeStartX = rz.Left + x
    m_ResizeOrigScale = m_Scale
    SetRect "kb_outline", m_BaseLeft, m_BaseTop + BASE_TITLEBAR_Y * m_Scale, BoardWidth(m_Scale), BoardHeightTotal(m_Scale)
    ShowOutline
End Sub

Public Sub OnResizeMouseMove(x As Single, y As Single)
    If Not m_ResizeArmed Then Exit Sub
    Dim newScale As Single: newScale = ComputeResizeScale(x)
    SetRect "kb_outline", m_BaseLeft, m_BaseTop + BASE_TITLEBAR_Y * newScale, BoardWidth(newScale), BoardHeightTotal(newScale)
End Sub

Public Sub OnResizeMouseUp(x As Single, y As Single)
    If Not m_ResizeArmed Then Exit Sub
    m_ResizeArmed = False
    m_Scale = ComputeResizeScale(x)
    HideOutline
    LayoutBoard
End Sub

Private Function ComputeResizeScale(x As Single) As Single
    Dim rz As Shape: Set rz = m_Ws.Shapes("kb_resize")
    Dim curX As Single: curX = rz.Left + x
    Dim dx As Single: dx = curX - m_ResizeStartX
    Dim newScale As Single
    newScale = m_ResizeOrigScale + dx / (2 * BASE_COL_GAP + BASE_COL_W)
    If newScale < 0.5 Then newScale = 0.5
    If newScale > 2 Then newScale = 2
    ComputeResizeScale = newScale
End Function

Private Sub BuildSummaryList()
    Dim hdrTitles(2) As String
    hdrTitles(0) = "タスク名": hdrTitles(1) = "期日": hdrTitles(2) = "達成状況"
    Dim hdrW(2) As Single
    hdrW(0) = 220: hdrW(1) = 90: hdrW(2) = 90

    Dim hx As Single: hx = m_BaseLeft
    Dim i As Integer
    For i = 0 To 2
        Dim shpH As Shape
        Set shpH = m_Ws.Shapes.AddShape(msoShapeRectangle, hx, m_BaseTop, hdrW(i), 24)
        shpH.Name = "kb_sumhdr" & i
        shpH.TextFrame.Characters().Text = hdrTitles(i)
        shpH.TextFrame.Characters().Font.Bold = True
        shpH.TextFrame.Characters().Font.Size = 9
        shpH.TextFrame.Characters().Font.Color = RGB(255, 255, 255)
        shpH.Fill.ForeColor.RGB = RGB(43, 63, 99)
        shpH.Line.Visible = False
        ' Excelのネイティブ Group() は再レイアウトのたびに内部の幅比率が
        ' 少しずつ崩れる不具合があるため使わない。3つを常に同じ基準位置から
        ' 再配置することでズレなく並べ、ドラッグでの分解防止はシート保護（Setup末尾）で行う。
        m_SummaryNames.Add shpH.Name
        hx = hx + hdrW(i)
    Next i

    Dim oleSum As OLEObject
    Set oleSum = m_Ws.OLEObjects.Add(ClassType:="Forms.ListBox.1", _
        Left:=m_BaseLeft, Top:=m_BaseTop + 24, Width:=400, Height:=200)
    oleSum.Name = "kb_sumlist"
    Set m_SummaryList = oleSum.Object
    m_SummaryList.Font.Size = 9
    m_SummaryList.BackColor = RGB(252, 254, 255)
    m_SummaryList.SpecialEffect = fmSpecialEffectSunken
    m_SummaryList.ColumnCount = 3
    m_SummaryList.ColumnWidths = "220 pt;90 pt;90 pt"
    m_SummaryNames.Add oleSum.Name

    Dim hSum As New KanbanHelper
    Set hSum.lb = m_SummaryList
    hSum.colIdx = -1   ' -1 = 簡易リスト（クリックでフルボードを開く）
    Set hSum.formRef = Me
    Set m_Helpers(m_HCount) = hSum
    m_HCount = m_HCount + 1
End Sub

Private Sub LoadSummary()
    Dim ws As Worksheet
    On Error GoTo NoSheet
    Set ws = ThisWorkbook.Sheets("タスクリスト")
    On Error GoTo 0
    m_SummaryList.Clear
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = "" Then Exit For
        m_SummaryList.AddItem CStr(ws.Cells(r, 2).Value)
        Dim ri As Long: ri = m_SummaryList.ListCount - 1
        Dim dv As Variant: dv = ws.Cells(r, 4).Value
        If IsDate(dv) Then m_SummaryList.List(ri, 1) = Format(CDate(dv), "yyyy/mm/dd") Else m_SummaryList.List(ri, 1) = CStr(dv)
        m_SummaryList.List(ri, 2) = CStr(ws.Cells(r, 6).Value)
    Next r
    Exit Sub
NoSheet:
    MsgBox "タスクリストシートが見つかりません。", vbCritical
End Sub

Private Sub ShowFullBoard()
    ' 設定ブロックの行を表示すると B11 の位置が下にずれるため、
    ' 行の表示切替後に基準位置を測り直してから再配置する
    ToggleSettingsRows True
    m_BaseTop = m_Ws.Range("B11").Top
    LayoutBoard
    ToggleVisibility m_SummaryNames, False
    ToggleVisibility m_BoardNames, True
    LoadFromExcel
End Sub

Private Sub ShowSummaryView()
    ToggleSettingsRows False
    m_BaseTop = m_Ws.Range("B11").Top
    LayoutSummary
    ToggleVisibility m_BoardNames, False
    ToggleVisibility m_SummaryNames, True
    LoadSummary
End Sub

Private Sub LayoutSummary()
    Dim hdrW(2) As Single
    hdrW(0) = 220: hdrW(1) = 90: hdrW(2) = 90
    Dim hx As Single: hx = m_BaseLeft
    Dim i As Integer
    For i = 0 To 2
        SetRect "kb_sumhdr" & i, hx, m_BaseTop, hdrW(i), 24
        hx = hx + hdrW(i)
    Next i
    SetRect "kb_sumlist", m_BaseLeft, m_BaseTop + 24, 400, 200
End Sub

Private Sub ClearExisting()
    Dim i As Long
    For i = m_Ws.OLEObjects.Count To 1 Step -1
        If Left(m_Ws.OLEObjects(i).Name, 3) = "kb_" Then m_Ws.OLEObjects(i).Delete
    Next i

    Dim toDelete As Collection
    Set toDelete = New Collection
    Dim s As Shape
    For Each s In m_Ws.Shapes
        If Left(s.Name, 3) = "kb_" Then toDelete.Add s
    Next s
    Dim sh As Shape
    For Each sh In toDelete
        sh.Delete
    Next sh
End Sub

Private Sub BuildAssigneePills()
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("名簿")
    On Error GoTo 0
    If wsC Is Nothing Then Exit Sub

    Dim shpHint As Shape
    Set shpHint = m_Ws.Shapes.AddShape(msoShapeRectangle, m_BaseLeft, m_BaseTop - 30, 230, 24)
    shpHint.Name = "kb_hint"
    shpHint.Fill.Visible = msoFalse
    shpHint.Line.Visible = msoFalse
    shpHint.TextFrame.Characters().Text = "担当者をドラッグ→列へドロップで新規タスク追加"
    shpHint.TextFrame.Characters().Font.Size = 8
    shpHint.TextFrame.Characters().Font.Bold = True
    shpHint.TextFrame.Characters().Font.Color = RGB(60, 60, 60)
    shpHint.TextFrame.HorizontalAlignment = xlHAlignLeft
    m_BoardNames.Add shpHint.Name

    m_PillCount = 0
    Dim lr As Long, r As Long, px As Single
    px = m_BaseLeft + 250
    lr = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lr
        If wsC.Cells(r, 1).Value = "" Then Exit For
        Dim nm As String: nm = CStr(wsC.Cells(r, 1).Value)
        Dim pw As Single: pw = 16 + Len(nm) * 7

        Dim olePill As OLEObject
        Set olePill = m_Ws.OLEObjects.Add(ClassType:="Forms.Label.1", _
            Left:=px, Top:=m_BaseTop - 30, Width:=pw, Height:=24)
        olePill.Name = "kb_pill" & m_PillCount
        Dim pill As MSForms.Label
        Set pill = olePill.Object
        pill.Caption = "  " & nm & "  "
        pill.BackColor = RGB(90, 110, 160)
        pill.ForeColor = RGB(255, 255, 255)
        pill.Font.Bold = True: pill.Font.Size = 9
        pill.TextAlign = fmTextAlignCenter
        pill.SpecialEffect = fmSpecialEffectRaised
        pill.MousePointer = fmMousePointerSizeAll
        Set m_Pills(m_PillCount) = pill
        m_BoardNames.Add olePill.Name
        px = px + pw + 8

        Dim hP As New KanbanHelper
        Set hP.pillLbl = pill
        hP.pillIndex = m_PillCount
        hP.pillName = nm
        Set hP.formRef = Me
        Set m_Helpers(m_HCount) = hP
        m_HCount = m_HCount + 1

        m_PillCount = m_PillCount + 1
    Next r
End Sub

Public Sub OnPillMouseDown(idx As Integer, nm As String, x As Single, y As Single)
    m_DragAssignee = nm
    Dim pill As MSForms.Label: Set pill = m_Pills(idx)
    lblGhost.Caption = "→ " & nm
    lblGhost.Width = 24 + Len(nm) * 8
    lblGhost.Left = pill.Left + x - lblGhost.Width / 2
    lblGhost.Top = pill.Top + y - lblGhost.Height / 2
    lblGhost.Visible = True
End Sub

Public Sub OnPillMouseMove(idx As Integer, x As Single, y As Single)
    If m_DragAssignee = "" Then Exit Sub
    Dim pill As MSForms.Label: Set pill = m_Pills(idx)
    lblGhost.Left = pill.Left + x - lblGhost.Width / 2
    lblGhost.Top = pill.Top + y - lblGhost.Height / 2
End Sub

Public Sub OnPillMouseUp(idx As Integer, x As Single, y As Single)
    If m_DragAssignee = "" Then Exit Sub
    Dim pill As MSForms.Label: Set pill = m_Pills(idx)
    Dim dropX As Single, dropY As Single
    dropX = pill.Left + x
    dropY = pill.Top + y

    Dim ci As Integer, target As Integer
    target = -1
    For ci = 0 To 2
        Dim lb As MSForms.ListBox: Set lb = m_Lists(ci)
        If dropX >= lb.Left And dropX <= lb.Left + lb.Width And _
           dropY >= lb.Top And dropY <= lb.Top + lb.Height Then
            target = ci
            Exit For
        End If
    Next ci

    Dim assignee As String: assignee = m_DragAssignee
    lblGhost.Visible = False
    m_DragAssignee = ""

    If target >= 0 Then QuickAddTask target, assignee
End Sub

Public Sub OnCardMouseDown(ci As Integer, x As Single, y As Single)
    Dim lb As MSForms.ListBox: Set lb = m_Lists(ci)
    If lb.ListIndex < 0 Then Exit Sub
    m_DragColIdx = ci
    m_DragItemIdx = lb.ListIndex
    m_DragArmed = True
    m_DragVisible = False
    m_DragStartX = lb.Left + x
    m_DragStartY = lb.Top + y
End Sub

Public Sub OnCardMouseMove(ci As Integer, x As Single, y As Single)
    If Not m_DragArmed Then Exit Sub
    Dim lb As MSForms.ListBox: Set lb = m_Lists(ci)
    Dim curX As Single, curY As Single
    curX = lb.Left + x: curY = lb.Top + y

    If Not m_DragVisible Then
        If Abs(curX - m_DragStartX) < 6 And Abs(curY - m_DragStartY) < 6 Then Exit Sub
        m_DragVisible = True
        lblGhost.Caption = "→ " & lb.List(m_DragItemIdx, 0)
        lblGhost.Width = 24 + Len(CStr(lb.List(m_DragItemIdx, 0))) * 7
        lblGhost.BackColor = RGB(40, 130, 90)
        lblGhost.Visible = True
    End If

    lblGhost.Left = curX - lblGhost.Width / 2
    lblGhost.Top = curY - lblGhost.Height / 2
End Sub

Public Sub OnCardMouseUp(ci As Integer, x As Single, y As Single)
    If Not m_DragArmed Then Exit Sub
    Dim wasVisible As Boolean: wasVisible = m_DragVisible
    Dim srcCol As Integer, srcIdx As Integer
    srcCol = m_DragColIdx: srcIdx = m_DragItemIdx

    m_DragArmed = False
    m_DragVisible = False
    m_DragColIdx = -1
    lblGhost.Visible = False

    If Not wasVisible Then Exit Sub  ' plain click, not a drag

    Dim lb As MSForms.ListBox: Set lb = m_Lists(ci)
    Dim dropX As Single, dropY As Single
    dropX = lb.Left + x: dropY = lb.Top + y

    Dim target As Integer, t As Integer
    target = -1
    For t = 0 To 2
        Dim tgt As MSForms.ListBox: Set tgt = m_Lists(t)
        If dropX >= tgt.Left And dropX <= tgt.Left + tgt.Width And _
           dropY >= tgt.Top And dropY <= tgt.Top + tgt.Height Then
            target = t
            Exit For
        End If
    Next t

    If target >= 0 And target <> srcCol Then MoveCardTo srcCol, srcIdx, target
End Sub

Private Sub MoveCardTo(fr As Integer, idxInSrc As Integer, to_ As Integer)
    Dim src As MSForms.ListBox: Set src = m_Lists(fr)
    If idxInSrc < 0 Or idxInSrc >= src.ListCount Then Exit Sub
    Dim dst As MSForms.ListBox: Set dst = m_Lists(to_)
    dst.AddItem src.List(idxInSrc, 0)
    Dim nr As Integer: nr = dst.ListCount - 1
    Dim c As Integer
    For c = 1 To 4
        dst.List(nr, c) = src.List(idxInSrc, c)
    Next c
    src.RemoveItem idxInSrc
    m_SelCol = -1: m_SelIdx = -1
    ResetHeaders
    SaveToExcel
End Sub

Private Sub QuickAddTask(ci As Integer, assignee As String)
    Dim n As String
    n = InputBox("タスク名:", "クイック追加（担当：" & assignee & " / " & ColName(ci) & "）")
    If n = "" Then Exit Sub
    Dim d As String: d = InputBox("期日（yyyy/mm/dd）:", "クイック追加", Format(Date + 7, "yyyy/mm/dd"))
    Dim r As String: r = InputBox("リマインド日数前:", "クイック追加", "3")
    Dim lb As MSForms.ListBox: Set lb = m_Lists(ci)
    lb.AddItem n
    Dim ri As Integer: ri = lb.ListCount - 1
    lb.List(ri, 1) = assignee
    lb.List(ri, 2) = d
    lb.List(ri, 3) = r
    lb.List(ri, 4) = "0"
    SaveToExcel
End Sub

Public Sub OnListClick(ci As Integer)
    If ci = -1 Then
        ShowFullBoard
        Exit Sub
    End If
    m_SelCol = ci
    m_SelIdx = m_Lists(ci).ListIndex
    Dim i As Integer
    Dim c(2) As Long
    c(0) = RGB(68, 114, 196): c(1) = RGB(220, 120, 0): c(2) = RGB(60, 150, 60)
    For i = 0 To 2
        m_Headers(i).Fill.ForeColor.RGB = c(i)
    Next i
    m_Headers(ci).Fill.ForeColor.RGB = RGB(20, 20, 20)
End Sub

Public Sub OnBtnClick(nm As String)
    Select Case nm
        Case "kb_btnA0": AddTask 0
        Case "kb_btnA1": AddTask 1
        Case "kb_btnA2": AddTask 2
        Case "kb_btnR0": MoveTask 0, 1
        Case "kb_btnL0": MoveTask 1, 0
        Case "kb_btnR1": MoveTask 1, 2
        Case "kb_btnL1": MoveTask 2, 1
        Case "kb_btnDel": DelTask
        Case "kb_btnClose": ShowSummaryView
    End Select
End Sub

Private Sub MoveTask(fr As Integer, to_ As Integer)
    If m_SelCol <> fr Or m_SelIdx < 0 Then
        MsgBox "移動元（" & ColName(fr) & "）のタスクを選択してください。", vbInformation
        Exit Sub
    End If
    MoveCardTo fr, m_SelIdx, to_
End Sub

Private Sub AddTask(ci As Integer)
    Dim n As String
    n = InputBox("タスク名:", "タスク追加（" & ColName(ci) & "）")
    If n = "" Then Exit Sub
    Dim a As String : a = InputBox("担当者名（名簿シートと一致）:", "タスク追加")
    Dim d As String : d = InputBox("期日（yyyy/mm/dd）:", "タスク追加", Format(Date + 7, "yyyy/mm/dd"))
    Dim r As String : r = InputBox("リマインド日数前:", "タスク追加", "3")
    Dim lb As MSForms.ListBox: Set lb = m_Lists(ci)
    lb.AddItem n
    Dim ri As Integer: ri = lb.ListCount - 1
    lb.List(ri, 1) = a
    lb.List(ri, 2) = d
    lb.List(ri, 3) = r
    lb.List(ri, 4) = "0"
    SaveToExcel
End Sub

Private Sub DelTask()
    If m_SelCol < 0 Or m_SelIdx < 0 Then
        MsgBox "削除するタスクを選択してください。", vbInformation: Exit Sub
    End If
    Dim lb As MSForms.ListBox: Set lb = m_Lists(m_SelCol)
    If lb.ListIndex < 0 Then Exit Sub
    If MsgBox("「" & lb.List(lb.ListIndex, 0) & "」を削除しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    lb.RemoveItem lb.ListIndex
    m_SelCol = -1: m_SelIdx = -1
    SaveToExcel
End Sub

Public Sub LoadFromExcel()
    Dim ws As Worksheet
    On Error GoTo NoSheet
    Set ws = ThisWorkbook.Sheets("タスクリスト")
    On Error GoTo 0
    Dim i As Integer
    For i = 0 To 2: m_Lists(i).Clear: Next i
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = "" Then Exit For
        Dim ci As Integer
        Select Case ws.Cells(r, 6).Value
            Case "完了":   ci = 2
            Case "進行中": ci = 1
            Case Else:     ci = 0
        End Select
        Dim lb As MSForms.ListBox: Set lb = m_Lists(ci)
        lb.AddItem CStr(ws.Cells(r, 2).Value)
        Dim ri As Integer: ri = lb.ListCount - 1
        lb.List(ri, 1) = CStr(ws.Cells(r, 3).Value)
        Dim dv As Variant: dv = ws.Cells(r, 4).Value
        If IsDate(dv) Then lb.List(ri, 2) = Format(CDate(dv), "yy/mm/dd") Else lb.List(ri, 2) = CStr(dv)
        lb.List(ri, 3) = CStr(ws.Cells(r, 5).Value)
        lb.List(ri, 4) = CStr(r)
    Next r
    Exit Sub
NoSheet:
    MsgBox "タスクリストシートが見つかりません。", vbCritical
End Sub

Private Sub SaveToExcel()
    Dim ws As Worksheet
    On Error GoTo NoSheet
    Set ws = ThisWorkbook.Sheets("タスクリスト")
    On Error GoTo 0
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lr >= 2 Then ws.Rows("2:" & lr).ClearContents
    Dim rn As Long : rn = 2
    Dim sn(2) As String
    sn(0) = "未完了" : sn(1) = "進行中" : sn(2) = "完了"
    Dim ci As Integer
    For ci = 0 To 2
        Dim lb As MSForms.ListBox: Set lb = m_Lists(ci)
        Dim i As Integer
        For i = 0 To lb.ListCount - 1
            ws.Cells(rn, 1).Value = rn - 1
            ws.Cells(rn, 2).Value = lb.List(i, 0)
            ws.Cells(rn, 3).Value = lb.List(i, 1)
            ws.Cells(rn, 4).Value = lb.List(i, 2)
            ws.Cells(rn, 5).Value = lb.List(i, 3)
            ws.Cells(rn, 6).Value = sn(ci)
            ws.Cells(rn, 7).Value = ""
            rn = rn + 1
        Next i
    Next ci
    Exit Sub
NoSheet:
    MsgBox "タスクリストシートが見つかりません。", vbCritical
End Sub

Private Sub WireBtn(b As MSForms.CommandButton)
    Dim h As New KanbanHelper
    Set h.btn = b
    Set h.formRef = Me
    Set m_Helpers(m_HCount) = h
    m_HCount = m_HCount + 1
End Sub

Private Sub AddArrow(cap As String, x As Single, y As Single, nm As String)
    Dim oleB As OLEObject
    Set oleB = m_Ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
        Left:=x, Top:=y, Width:=20, Height:=22)
    oleB.Name = nm
    Dim b As MSForms.CommandButton
    Set b = oleB.Object
    b.Caption = cap
    b.Font.Bold = True: b.Font.Size = 9
    WireBtn b
    m_BoardNames.Add oleB.Name
End Sub

Private Sub ResetHeaders()
    m_Headers(0).Fill.ForeColor.RGB = RGB(68, 114, 196)
    m_Headers(1).Fill.ForeColor.RGB = RGB(220, 120, 0)
    m_Headers(2).Fill.ForeColor.RGB = RGB(60, 150, 60)
End Sub

Private Function ColName(i As Integer) As String
    Select Case i
        Case 0: ColName = "未完了"
        Case 1: ColName = "進行中"
        Case 2: ColName = "完了"
    End Select
End Function
