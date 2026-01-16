Attribute VB_Name = "m9_Автозапуск_по_времени"
Sub АвтозапускПоВремени()
If MsgBox("Запустить старт по времени: CallКалькулятор, CallЗаполнитьОтчетПоСканированиюФКБ?", vbYesNo) = vbNo Then Exit Sub

        Application.OnTime TimeValue("15:27:00"), "CallКалькулятор"
            
        Application.OnTime TimeValue("15:26:00"), "CallRidNacs"

        Application.OnTime TimeValue("15:25:00"), "CallпраздникиСегодня"
End Sub


Sub CallКалькулятор() ' Для горячей клавиши и / или вызова программы
    shell "C:\Windows\System32\calc.exe", vbNormalFocus
End Sub


Sub CallЗаполнитьОтчетПоСканированиюФКБ() ' Для горячей клавиши и / или вызова программы
    Call ЗаполнитьОтчетПоСканированиюФКБ(Nothing)
End Sub


Sub CallRidNacs() ' Для горячей клавиши и / или вызова программы
    shell "C:\Users\s.lazarev\Documents\DISTRIBUTIVE\RidNacs.exe", vbNormalFocus
End Sub


Sub CallпраздникиСегодня() ' Для горячей клавиши и / или вызова программы
    Call праздникиСегодня(Nothing)
End Sub



Sub АвтозапускНаВыходные(control As IRibbonControl) 'Если текущ дата больше 13 июля 25 г, то старт Автозаполн Отчет по фабрике и старт БТР. Иначе - запись состояния в лог.
    If MsgBox("Запустить старт по времени: 777, 888?", vbYesNo) = vbNo Then Exit Sub
    
            Application.OnTime TimeValue("09:00:00"), "Запись_в_лог" ' 9 ч 12.07.2025 Сб
            Application.Wait Now + TimeValue("00:00:05")
            Application.OnTime TimeValue("05:30:00"), "ПроверкаЦелевыхДатыВремени"
            Application.Wait Now + TimeValue("00:00:05")
            Application.OnTime TimeValue("20:00:00"), "Запись_в_лог" ' 20 ч 12.07.2025 Сб
            Application.Wait Now + TimeValue("00:00:05")
            
            Application.OnTime TimeValue("09:00:00"), "Запись_в_лог" ' 9 ч 13.07.2025 Вс
            Application.Wait Now + TimeValue("00:00:05")
            Application.OnTime TimeValue("20:00:00"), "Запись_в_лог" ' 20 ч 13.07.2025 Вс
            Application.Wait Now + TimeValue("00:00:05")
            Application.OnTime TimeValue("05:00:00"), "ПроверкаЦелевыхДатыВремени" ' 05 ч 14.07.2025 ПН

     
End Sub

Sub ПроверкаЦелевыхДатыВремени() ' Для горячей клавиши и / или вызова программы
    Dim strFile_Path As String
        strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
    Dim targetDate, targetTime, currentDate, currentTime As Date
                    targetDate = DateValue("2025-07-13") ' Замените на вашу назначенную дату
                    targetTime = TimeValue("23:59:00") ' Замените на ваше назначенное время
                    currentDate = Date ' Получите текущее время и дату
                    currentTime = Time
    If currentDate > targetDate Or (currentDate = targetDate And currentTime > targetTime) Then ' Проверьте условия
       
        Open strFile_Path For Append As #1
        Print #1, " "
        Print #1, "Отработал Саб ПроверкаЦелевыхДатыВремени, целевая Дата(время) НАСТУПИЛА, " & Now
        Close #1
        Application.OnTime TimeValue("05:42:00"), "ЗаполнитьОтчетФабрика"
        Application.OnTime TimeValue("06:25:00"), "БитриксНачало"
    Else
        Open strFile_Path For Append As #1
        Print #1, " "
        Print #1, "Отработал Саб ПроверкаЦелевыхДатыВремени, целевая Дата(время) 2025-07-13 23:59:00 не наступила. " & Now
        Close #1
        Call ParserБитрикс24
    End If
End Sub

Sub Запись_в_лог() ' Для записи состояния в лог.
    Dim strFile_Path As String
        username = Environ("UserName")
        strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
    Open strFile_Path For Append As #1
    Print #1, " "
    Print #1, "Отработал Саб Запись_в_лог, Эксель норм, " & username & " " & Now   '
    Close #1
    Call ParserБитрикс24
End Sub



Sub ПроверкаНаПонедельник14Июля()
    Dim strFile_Path As String
        strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
    If Date = "14.07.2025" Then
        Open strFile_Path For Append As #1
        Print #1, " "
        Print #1, "Отработал Саб ПроверкаНаПонедельник14Июля, Так вот, Понедельник14 Июля - НАСТУПИЛ и будут запущены Сабы: ЗаполнитьОтчетФабрика и БитриксНачало  " & Now
        Close #1
        Application.OnTime TimeValue("05:42:00"), "ЗаполнитьОтчетФабрика"
        Application.OnTime TimeValue("06:25:00"), "БитриксНачало"
    Else
        Open strFile_Path For Append As #1
        Print #1, " "
        Print #1, "Отработал Саб ПроверкаНаПонедельник14Июля, Понедельник14 Июля НЕ наступил, " & Now
        Close #1
    End If
End Sub


Sub ПроверкаЦелевыхДатыВремени_2()
If Date = "12.07.2025" Then
MsgBox "Верно, сегодня 11.07.2025"
Else
MsgBox "Сегодня НЕ 12.07.2025!"
End If
End Sub
