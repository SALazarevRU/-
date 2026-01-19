VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f1_ТаймерФКБ 
   Caption         =   "Работает сканер..."
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16260
   OleObjectBlob   =   "f1_ТаймерФКБ.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "f1_ТаймерФКБ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
            CommandButton1.BackColor = RGB(204, 58, 0)
            CommandButton1.Font.Bold = True
            CommandButton1.caption = "ДОП. ВРЕМЯ 7 СЕК"
            CommandButton1.ForeColor = RGB(225, 225, 225)
                Application.Wait Now + TimeValue("00:00:07")
            CommandButton1.BackColor = RGB(190, 190, 190)
            
            CommandButton1.ForeColor = RGB(0, 0, 0)
            CommandButton1.Font.Bold = False
End Sub

Private Sub CommandButton2_Click()
            CommandButton2.BackColor = RGB(204, 58, 0)
                Application.Wait Now + TimeValue("00:00:14")
            CommandButton2.BackColor = RGB(190, 190, 190)
End Sub

Private Sub CommandButton3_Click()
            CommandButton3.BackColor = RGB(204, 58, 0)  'IIf(CommandButton1.BackColor = -2147483633, vbGreen, -2147483633)
                Application.Wait Now + TimeValue("00:00:40")
            CommandButton3.BackColor = RGB(190, 190, 190)
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub UserForm_Initialize()
f1_ТаймерФКБ.Label8.caption = Format(Now, "dd MM yyyy  HH:mm:ss")

    Me.StartUpPosition = 0 'Моя стартовая позиция
        Me.Top = 290 + Application.Top
        Me.Left = 380 + Application.Left
        
'    iTimer = TimeValue("00:00:10")
    On Error GoTo NoTimerValue
    iTimer888 = TimeSerial(0, 0, CLng(Workbooks("Итог_ФКБ_Лазарев.xlsm").Worksheets("ппонФКБ").Range("D41").Value))   ' Укажите вместо "Sheet1" актуальное имя вашего листа
    

    Call TimerStartФКБ
    Exit Sub
    
NoTimerValue:
    MsgBox "Ошибка чтения времени из ячейки AZ3. Проверьте, что там число а не текст или какие-то спец символы! ", vbCritical
    Unload Me
    
    
End Sub
