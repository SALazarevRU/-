VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
 Private WithEvents app As Application
Attribute app.VB_VarHelpID = -1
 
 Private Sub Workbook_Open()
'    Worksheets("Расширенный реестр").Range("AX1").Value = 8
'   Worksheets("Расширенный реестр").Range("AX1").Value = "ВремяСканирования"
'   Call CallВремяСканирования
 
 
''        Application.OnKey "{RIGHT}", "CallОткрытьПапкуДинамика"   ' Перехват клавиши  RIGHT - ВПРАВО.
        Application.OnKey "{UP}", "CallСканингГПБ"   ' Перехват клавиши  UP - ВВЕРХ. "CallTest" - главная процедура запуска Автовалидации.
''        Application.OnKey "{F12}", "CallОткрытьФайлДинамика"   ' Перехват клавиши  UP - ВВЕРХ.
''       Call CallTVozduha
      
        Set app = Application
        Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev"
            username = Environ("UserName")  ' Получаем имя пользователя.
               If username <> SpecifiedUserName Then
                 Dim targetDate, targetTime, currentDate, currentTime As Date
                    targetDate = DateValue("2026-11-15") ' Замените на вашу назначенную дату
                    targetTime = TimeValue("03:00:00") ' Замените на ваше назначенное время
                   
                    ' Получите текущее время и дату
                    currentDate = Date
                    currentTime = Time
                    
                    If currentDate > targetDate Or (currentDate = targetDate And currentTime > targetTime) Then ' Проверьте условия
                       
            '                Call ОтключитьНадстройку
                           DeleteMyAddIn  ' Удалить надстройку.
                    End If ' вот это тут походу лишнее !!!!!!
                  Else
                    Exit Sub
               End If
End Sub
 
 Private Sub Workbook_BeforeClose(Cancel As Boolean)
 On Error Resume Next
'Call CallБэкапЭтойКниги
' Call БэкапЭтойКниги(Nothing)
 Call ЗаписатьВЛогCLOSE
'    Application.DisplayAlerts = False
'    ThisWorkbook.Save ' это надстройка, а актив.воркбук это .xlsx
'    Application.DisplayAlerts = True
    
    Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev"
        username = Environ("UserName")  ' Получаем имя пользователя.
           If username <> SpecifiedUserName Then
    
                Dim targetDate, targetTime, currentDate, currentTime As Date
                    targetDate = DateValue("2026-09-15") ' Замените на вашу назначенную дату
                    targetTime = TimeValue("03:00:00") ' Замените на ваше назначенное время
                    currentDate = Date ' Получите текущее время и дату
                    currentTime = Time
                If currentDate > targetDate Or (currentDate = targetDate And currentTime > targetTime) Then ' Проверьте условия
                        Call ОтключитьНадстройку
                End If ' вот это тут походу лишнее !!!!!!
           Else  ' Иначе:
              Exit Sub  ' Остановка кода.
           End If
End Sub
