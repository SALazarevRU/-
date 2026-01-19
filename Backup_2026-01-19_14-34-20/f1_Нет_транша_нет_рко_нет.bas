VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f1_Нет_транша_нет_рко_нет 
   Caption         =   "© Клавиатура выбора Комментария"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "f1_Нет_транша_нет_рко_нет.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f1_Нет_транша_нет_рко_нет"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()
            activeCell.Value = "нет транша"
            Unload Me
End Sub

Private Sub CommandButton4_Click()
            activeCell.Value = "нет рко/рнко"
            Unload Me
End Sub

Private Sub CommandButton5_Click()
            activeCell.Value = "нет"
            Unload Me
End Sub

Private Sub CommandButton6_Click()
            activeCell.Value = "нет транша, нет рко/рнко"
            Unload Me
End Sub

Private Sub CommandButton7_Click()
            Unload Me
End Sub


Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = 320 + Application.Top
    Me.Left = 370 + Application.Left
End Sub
