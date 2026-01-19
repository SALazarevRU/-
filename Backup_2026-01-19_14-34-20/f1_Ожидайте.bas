VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f1_Ожидайте 
   Caption         =   "Просто подождите завершения процесса :)"
   ClientHeight    =   9480.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13230
   OleObjectBlob   =   "f1_Ожидайте.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f1_Ожидайте"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
Dim info As String
'info = "Дата:" & vbCrLf
'f1_Ожидайте.Label2.Caption = info & Format(Now, "dddddd hh ч. mm мин")

f1_Ожидайте.Label2.caption = Format(Now, "dd MM yyyy  HH:mm:ss")


End Sub
