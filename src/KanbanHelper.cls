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

Public WithEvents lb  As MSForms.ListBox
Public WithEvents btn As MSForms.CommandButton
Public colIdx As Integer
Public formRef As Object

Private Sub lb_Click()
    If Not formRef Is Nothing Then formRef.OnListClick colIdx
End Sub

Private Sub btn_Click()
    If Not formRef Is Nothing Then formRef.OnBtnClick btn.Name
End Sub
