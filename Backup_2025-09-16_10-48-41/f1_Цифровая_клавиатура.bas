VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f1_Цифровая_клавиатура 
   Caption         =   "© Цифровая клавиатура"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2985
   OleObjectBlob   =   "f1_Цифровая_клавиатура.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f1_Цифровая_клавиатура"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
ActiveCell.Value = 3
End Sub

Private Sub CommandButton11_Click()
Unload Me
End Sub

Private Sub CommandButton12_Click()
ActiveCell.ClearContents
End Sub

Private Sub CommandButton2_Click()
If ActiveCell.Value = "" Then
   ActiveCell.Value = 36
Else
ActiveCell.Value = ActiveCell.Value & "6"
End If
End Sub

Private Sub CommandButton3_Click()
    If ActiveCell.Value = "" Then
       ActiveCell.Value = 1
    Else
        ActiveCell.Value = ActiveCell.Value & "1"
    End If
End Sub

Private Sub CommandButton4_Click()
ActiveCell.Value = 2
End Sub

Private Sub CommandButton6_Click()
    If ActiveCell.Value = "" Then
       ActiveCell.Value = 5
    Else
        ActiveCell.Value = ActiveCell.Value & "5"
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = 290 + Application.Top
    Me.Left = 550 + Application.Left
End Sub
