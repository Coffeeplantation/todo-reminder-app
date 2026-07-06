VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "KanbanHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents lb      As MSForms.ListBox
Public WithEvents btn     As MSForms.CommandButton
Public WithEvents pillLbl As MSForms.Label
Public colIdx As Integer
Public formRef As Object
Public pillIndex As Integer
Public pillName As String

Private Sub lb_Click()
    If Not formRef Is Nothing Then formRef.OnListClick colIdx
End Sub

Private Sub lb_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not formRef Is Nothing Then formRef.OnCardMouseDown colIdx, X, Y
End Sub

Private Sub lb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not formRef Is Nothing Then formRef.OnCardMouseMove colIdx, X, Y
End Sub

Private Sub lb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not formRef Is Nothing Then formRef.OnCardMouseUp colIdx, X, Y
End Sub

Private Sub btn_Click()
    If Not formRef Is Nothing Then formRef.OnBtnClick btn.Name
End Sub

Private Sub pillLbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not formRef Is Nothing Then formRef.OnPillMouseDown pillIndex, pillName, X, Y
End Sub

Private Sub pillLbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not formRef Is Nothing Then formRef.OnPillMouseMove pillIndex, X, Y
End Sub

Private Sub pillLbl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not formRef Is Nothing Then formRef.OnPillMouseUp pillIndex, X, Y
End Sub
