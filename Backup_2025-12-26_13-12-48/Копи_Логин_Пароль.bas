VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Копи_Логин_Пароль 
   Caption         =   "Данные для авторизации"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   OleObjectBlob   =   "Копи_Логин_Пароль.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Копи_Логин_Пароль"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
       .SetText Логин: .PutInClipboard
    End With
End Sub

Private Sub CommandButton2_Click()
    Application.CutCopyMode = False
        With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
                .SetText ПарольПО: .PutInClipboard
        End With
        Start = Timer  ' Определяем время старта
        Do While Timer < Start + 2
            DoEvents  ' Уступаем другим процессам
        Loop
    Unload Копи_Логин_Пароль
  
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    With Me 'Перед закрытием запомнить позицию
        SaveSetting "Ms Office", "Данные для авторизации", "Left", .Left
        SaveSetting "Ms Office", "Данные для авторизации", "Top", .Top
    End With
End Sub
Private Sub UserForm_Initialize()
    With Me
        If Application.Left > -100 Then
            .StartUpPosition = 0
            .Left = GetSetting("Ms Office", "Данные для авторизации", "Left", .Left)
            .Top = GetSetting("Ms Office", "Данные для авторизации", "Top", .Top)
            If .Left <= 0 Or .Left > (Application.Left + Application.Width - 900) Or _
            .Top <= 0 Or .Top > (Application.Top + Application.Height - 900) Then
                'Если сохраненная ранее позиция вышла за предел экрана
                .StartUpPosition = 2
            End If
        End If
    End With
End Sub
