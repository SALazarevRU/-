VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFindBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cb As Office.CommandBar

Private Sub Class_Initialize()
    Set cb = Application.CommandBars.Add(, msoBarPopup, , True)
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    cb.Delete
    Set cb = Nothing
End Sub

Public Sub Show()
    cb.ShowPopup
End Sub

Public Sub Add(rng As Range)
    Dim cbb As Office.CommandBarButton
    Set cbb = cb.Controls.Add(1&, , , , True)
    With cbb
        .caption = Left$(rng.Value, 64)
        .Style = msoButtonCaption
        .Parameter = rng.Parent.Name
        .Tag = rng.Address
        .OnAction = "=SelectCell()"
    End With
End Sub

Public Property Get Count() As Integer
    Count = cb.Controls.Count
End Property

