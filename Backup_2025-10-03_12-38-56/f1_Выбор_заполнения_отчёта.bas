VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f1_Выбор_заполнения_отчёта 
   Caption         =   "Выбор метода заполнения отчёта"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "f1_Выбор_заполнения_отчёта.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f1_Выбор_заполнения_отчёта"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    f1_Выбор_заполнения_отчёта.Label1.Caption = Format(Now, "dd MMMM yyyy  HH:mm")
    CommandButton1.Caption = "Тихое заполнение" & vbNewLine & "(без открытия файла)"

    Me.StartUpPosition = 0 'Моя стартовая позиция
        Me.Top = 290 + Application.Top
        Me.Left = 580 + Application.Left
End Sub
