VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f1_ТаймерГПБ 
   Caption         =   "Работает сканер..."
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17280
   OleObjectBlob   =   "f1_ТаймерГПБ.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "f1_ТаймерГПБ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
            CommandButton1.BackColor = RGB(255, 255, 0)
            CommandButton1.Font.Bold = True
            CommandButton1.Caption = "ДОПОЛНИТЕЛЬНОЕ ВРЕМЯ 7 СЕК"
            CommandButton1.ForeColor = RGB(0, 0, 0) ' цвет текста в элементе управления (0, 0, 0) -черный
'                Application.Wait Now + TimeValue("00:00:07")
                 Application.Wait Now + TimeSerial(0, 0, 7)
            CommandButton1.BackColor = RGB(190, 190, 190)
            CommandButton1.Caption = "Доп.время"
            CommandButton1.ForeColor = RGB(255, 0, 50) ' цвет текста в элементе управления красный
            CommandButton1.Font.Bold = False
End Sub

Private Sub CommandButton2_Click()
            CommandButton2.BackColor = RGB(255, 255, 0) 'RGB(204, 58, 0)красный
            CommandButton2.Font.Bold = True
            CommandButton2.Caption = "ДОПОЛНИТЕЛЬНОЕ ВРЕМЯ 14 СЕК"
            CommandButton2.ForeColor = RGB(0, 0, 0)
                Application.Wait Now + TimeValue("00:00:14")
            CommandButton2.BackColor = RGB(190, 190, 190)
            CommandButton2.Caption = "Доп.время"
            CommandButton2.ForeColor = RGB(0, 0, 0)
            CommandButton2.Font.Bold = False
End Sub

Private Sub CommandButton3_Click()
            CommandButton3.BackColor = RGB(255, 255, 0)  'IIf(CommandButton1.BackColor = -2147483633, vbGreen, -2147483633)
            CommandButton3.Font.Bold = True
            CommandButton3.Caption = "ДОПОЛНИТЕЛЬНОЕ ВРЕМЯ 40 СЕК"
            CommandButton3.ForeColor = RGB(0, 0, 0)
                Application.Wait Now + TimeValue("00:00:40")
            CommandButton3.BackColor = RGB(190, 190, 190)
            CommandButton3.Caption = "Доп.время"
            CommandButton3.ForeColor = RGB(0, 0, 0)
            CommandButton3.Font.Bold = False
End Sub

Private Sub CommandButton4_Click()
            CommandButton4.BackColor = RGB(255, 255, 0)  'IIf(CommandButton1.BackColor = -2147483633, vbGreen, -2147483633)
            CommandButton4.Font.Bold = True
            CommandButton4.Caption = "60 сек"
            CommandButton4.ForeColor = RGB(0, 0, 0)
                Application.Wait Now + TimeValue("00:01:00")
            CommandButton4.BackColor = RGB(190, 190, 190)
            CommandButton4.Caption = "Доп.время"
            CommandButton4.ForeColor = RGB(0, 0, 0)
            CommandButton4.Font.Bold = False
End Sub

Private Sub CommandButton5_Click()
            CommandButton5.BackColor = RGB(255, 255, 0)  'IIf(CommandButton1.BackColor = -2147483633, vbGreen, -2147483633)
            CommandButton5.Font.Bold = True
            CommandButton5.Caption = "90 сек"
            CommandButton5.ForeColor = RGB(0, 0, 0)
                Application.Wait Now + TimeValue("00:01:30")
            CommandButton5.BackColor = RGB(190, 190, 190)
            CommandButton5.Caption = "Доп.время"
            CommandButton5.ForeColor = RGB(0, 0, 0)
            CommandButton5.Font.Bold = False
End Sub

Private Sub CommandButton6_Click()
            CommandButton6.BackColor = RGB(255, 255, 0)  'IIf(CommandButton1.BackColor = -2147483633, vbGreen, -2147483633)
            CommandButton6.Font.Bold = True
            CommandButton6.Caption = "110 сек"
            CommandButton6.ForeColor = RGB(0, 0, 0)
                Application.Wait Now + TimeValue("00:01:50")
            CommandButton6.BackColor = RGB(190, 190, 190)
            CommandButton6.Caption = "Доп.время"
            CommandButton6.ForeColor = RGB(0, 0, 0)
            CommandButton6.Font.Bold = False
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub UserForm_Initialize()
 
    Me.StartUpPosition = 0
    Me.Top = 270 + Application.Top
    Me.Left = 450 + Application.Left
 
'        iTimer = TimeValue("00:00:05")
 
    On Error GoTo NoTimerValue
    iTimer = TimeSerial(0, 0, CLng(Workbooks("03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.xlsx").Worksheets("Расширенный реестр").Range("AX1").Value))   ' Укажите вместо "Sheet1" актуальное имя вашего листа

    Call TimerStartГПБ
    Exit Sub
    
NoTimerValue:
    MsgBox "Ошибка чтения времени из ячейки AX1. Проверьте, что там число а не текст или какие-то спец символы! ", vbCritical
    Unload Me

End Sub
    
    
    
    
    
    
    
    

