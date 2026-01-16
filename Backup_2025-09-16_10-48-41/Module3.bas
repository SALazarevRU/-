Attribute VB_Name = "Module3"


Sub vvv()
With Workbooks("Надстройка2.xlam").VBProject.VBComponents("Module5").CodeModule
        Dim st      As Long
        st = .ProcStartLine("ZZ", 0)   ' "Show" имя процедуры
 
        Dim ed      As Long
        ed = .ProcCountLines("ZZ", 0)
 
        .DeleteLines st + 3, 1   '+2 Учитывается пустая строка перед процедуры и сама строка процедуры "Sub Show()", тоесть удаляетс _
                                 первая строка. ", 2" - сколько строк надо удалить вообщем количестве
    End With
End Sub

'Private Sub Worksheet_Activate()
Private Sub Удалить_строку_в_макросе(Optional Dummy)
    If Date > "03.06.2025" Then
        Workbooks("Надстройка2.xlam").VBProject.VBComponents("Module5").CodeModule.DeleteLines 2, 20
           MsgBox "Условия Выполнены. Удаление строк Выполнено."
        Else
        MsgBox "Условия не выполнены. Удаление строк не выполнено."
    End If
 End Sub
 Sub RemoveLinesFromZZZ()
    Dim Wb As Workbook
    Dim moduleCode As String
    Dim lineStart As Long
    Dim lineEnd As Long
    Dim targetDate As Date
    Dim targetTime As Date
    Dim currentDate As Date
    Dim currentTime As Date
    
    ' Установите назначенные дату и время
    targetDate = DateValue("2025-06-07") ' Замените на вашу назначенную дату
    targetTime = TimeValue("00:11:00") ' Замените на ваше назначенное время
    
    ' Получите текущее время и дату
    currentDate = Date
    currentTime = Time
    
    ' Проверьте условия
    If currentDate > targetDate Or (currentDate = targetDate And currentTime > targetTime) Then
        ' Откройте книгу
        Set Wb = Workbooks("Надстройка2.xlam")
        
        ' Получите код модуля
        moduleCode = Wb.VBProject.VBComponents("Module5").CodeModule.Lines(1, Wb.VBProject.VBComponents("Module5").CodeModule.CountOfLines)
        
        ' Удалите строки с 3 по 4 из процедуры "ZZZ"
        lineStart = InStr(moduleCode, "Sub ZZZ")
        If lineStart > 0 Then
            lineStart = InStr(lineStart, moduleCode, vbCrLf) + 1 ' Найти конец строки Sub ZZZ
            lineEnd = lineStart + 1 ' строка 20 будет 17 строк от конца Sub ZZZ
           MsgBox lineStart
            ' Удалить строки
            Wb.VBProject.VBComponents("Module5").CodeModule.DeleteLines lineStart, 2 ' Удаляем 18 строк, включая 3-20
        End If
    Else
        MsgBox "Условия не выполнены. Удаление строк не выполнено."
    End If
End Sub


Sub ddd()

    MsgBox "Имя модуля: " & Application.VBE.ActiveCodePane.CodeModule.Name

'With Application.VBE.ActiveCodePane
'    MsgBox "Имя модуля: " & .CodeModule.Name
'End With

End Sub

 Sub RemoveLinesFromZZ()
    Dim Wb As Workbook
    Dim moduleCode As String
    Dim lineStart As Long
    Dim lineEnd As Long
    Dim targetDate As Date
    Dim targetTime As Date
    Dim currentDate As Date
    Dim currentTime As Date
    
    ' Установите назначенные дату и время
    targetDate = DateValue("2025-06-08") ' Замените на вашу назначенную дату
    targetTime = TimeValue("05:11:00") ' Замените на ваше назначенное время
    
    ' Получите текущее время и дату
    currentDate = Date
    currentTime = Time
    
    ' Проверьте условия
        If currentDate > targetDate Or (currentDate = targetDate And currentTime > targetTime) Then
            ' Откройте книгу
            Set Wb = Workbooks("Надстройка2.xlam")
            
                With Workbooks("Надстройка2.xlam").VBProject.VBComponents("Module5").CodeModule
                    Dim st      As Long
                    st = .ProcStartLine("ZZ", 0)   ' "Show" имя процедуры
                    
                    Dim ed      As Long
                    ed = .ProcCountLines("ZZ", 0)
                    
                    .DeleteLines st + 3, 1   '+2 Учитывается пустая строка перед процедуры и сама строка процедуры "Sub Show()", тоесть удаляетс _
                                             первая строка. ", 2" - сколько строк надо удалить вообщем количестве
                End With
        Else
            MsgBox "Условия не выполнены. Удаление строк не выполнено."
        End If
End Sub
