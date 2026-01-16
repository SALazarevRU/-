VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f1_‘¿…À_—Œ’–¿Õ≈Õ 
   Caption         =   "© —ÓÓ·˘ÂÌËÂ Ó ÒÓı‡ÌÂÌËË Ù‡ÈÎ‡"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5610
   OleObjectBlob   =   "f1_‘¿…À_—Œ’–¿Õ≈Õ.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f1_‘¿…À_—Œ’–¿Õ≈Õ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ToggleButton1_Click()
            Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub ToggleButton2_Click()
   Unload Me
End Sub


Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = 290 + Application.Top
    Me.Left = 550 + Application.Left

f1_‘¿…À_—Œ’–¿Õ≈Õ.Label2.Caption = Format(Now, "  dd MMMM yyyy                                                                        HH:mm:ss")


End Sub
