Attribute VB_Name = "m9_Напоминание"
Public ТекстНапоминания As String 'для получения выбранного значения
Public времяНапоминания As Date 'для получения выбранного значения

'Sub Напоминание()
Public Sub Напоминание(control As IRibbonControl)
    
     ТекстНапоминания = InputBox("введите текст напоминания", "Запрос данных", " ")
  
    If ТекстНапоминания = "" Then
        MsgBox "Текст не введен, закрытие программы", vbCritical, "Текст напоминания пользователя"
        Exit Sub
    End If
    
    ' Запросить у пользователя ввод времени
    vRetValВремя = InputBox("Введите время напоминания в формате ЧЧ:ММ", "Напоминание о времени")
    
    ' Преобразовать ввод в формат времени
    On Error Resume Next
    времяНапоминания = TimeValue(vRetValВремя)
    On Error GoTo 0
    
    ' Проверка, если время введено неправильно
    If времяНапоминания = 0 Then
        MsgBox "Неверный формат времени. Программа закрыта."
        Exit Sub
    End If
    
    ' Запланировать выполнение процедуры ВремяОбеда
    Application.OnTime времяНапоминания, "ТекстНапоминанияПоказ"
    MsgBox "Напоминание запланировано на " & Format(времяНапоминания, "hh:mm"), vbYesNo, "Время показа запланированного напоминания"
End Sub

Sub ТекстНапоминанияПоказ()
    MsgBox ТекстНапоминания, vbYesNo, "Программа запланированных напоминаний"
End Sub

''Private Sub Workbook_BeforeClose(Cancel As Boolean)
''   Application.OnTime dTime, "Напоминание", , False
''End Sub
''Private Sub Workbook_Open()
''  Application.OnTime Now + TimeValue("00:15:00"), "Напоминание"
''End Sub
