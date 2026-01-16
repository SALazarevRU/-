Attribute VB_Name = "m9_ФКБ_Local_поиск_и_др"
'------------------------------------------------------------------------------------------------------------------------------
'Author    :
'DateTime  :
'Purpose   : Процедура поиска значения активной ячейки книги в строке второй закрытой книги на сервере [Итог.xlsx]
'            с проверкой на занятость книги, её открытием, сбросом фильтров, выделением строки с искомым значением, выбором варианта
'            действий.
'Status    : Процедура в рабочем состоянии, докуметирована
'-------------------------------------------------------------------------------------------------------------------------------


Sub Поиск_в_ФКБ_на_ПК(control As IRibbonControl)
    Call Поиск_в_ФКБ_Localniy_Ручник ' Передаем управление базовой процедуре
End Sub

Sub Поиск_в_ФКБ_Localniy_Ручник()
'    задаю переменные:
        Dim GCell As Range
        Dim txt$
        Dim wBook As Workbook
        Dim fl As Boolean
        
     Application.ScreenUpdating = False        'оператор VBA для отключения обновления экрана (=ускорение), он же фоновый режим(?))
        
     Cells(ActiveCell.Row, 2).Select 'этот код отбросит в ячейку с ClaimID с любой ячейке строки
'    подключаю функцию проверки IsBookOpen("wbFullName")на открытость/закрытость книги2 [Итог.xlsx]:
'    fl = IsBookOpen("C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ_NEW.xlsx")
'
'    MsgBox "Сергей Александрович!" & _
'        vbNewLine & "файл " & "Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ_NEW.xlsx'" & IIf(fl, " уже открыт", " не занят")
''    потом доделать: Если книга уже открыта, то узнать кем и отправить в чат или на почту сообщ-е "***, ***, плиз" + Exit Sub  (В реж.чтения не надо)
'    txt = ActiveCell.Value 'Искомое значение-значение активной ячейки активного листа активной книги1 [Шапка_6.xlsm] его и буду искать в книге2.
'    Debug.Print "ищу= "; txt 'Ок,выводит.
'    If Len(Trim(txt)) = 0 Then MsgBox "Ничего не выделено!", vbCritical: Exit Sub
    
    If fl Then ' если уже был открыт мною
    Set wBook = Workbooks("Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ_NEW.xlsx")
    wBook.Windows(1).Activate 'вытягиваем на первый план
    Else
    Set wBook = Workbooks.Open(FileName:="C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ_NEW.xlsx")  ' открыл
    End If
    
    With Application  ' размещаю окно  в координатах:
        .WindowState = xlNormal
        .Width = 1420 ' ШИРИНА окна
        .Height = 307 ' ВЫСОТА окна
        .Left = 0     ' ЛЕВАЯ точка
        .Top = 230      ' ВЕРХНЯЯ точка
    End With
    
    Application.ScreenUpdating = True        'включаю обновление экрана.
    
'   снимаю фильтры на активном листе [Итог.xlsx]:
    If wBook.ActiveSheet.FilterMode = True Then
         wBook.ActiveSheet.ShowAllData
    End If
 
        Set GCell = wBook.Sheets("Лист1").UsedRange.Columns(2).Find(What:=txt, LookIn:=xlValues, LookAt:=xlWhole)    ' произвожу в ней поиск полного совпадения
 
    If GCell Is Nothing Then      ' проверяю, является ли переменная GCell пустой
        MsgBox "ID " & txt & " не найден.", vbOKOnly + vbCritical, "РЕЗУЛЬТАТ ПОИСКА В ФАЙЛЕ ИТОГ"
        Workbooks("Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ.xlsx").Close SaveChanges:=False 'Чтобы закрыть книгу в Excel без сохранения изменений
    Else
        Workbooks("Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ_NEW.xlsx").Activate   ' на передний план [Итог.xlsx]
        ActiveSheet.Range(GCell, GCell.Offset(0, 35)).Select    ' выделяю найденный результат
        Debug.Print "GCell= " & GCell
    End If
    
    result = MsgBox("Вы хотите зарегистрировать документы в файле?" _
        & vbNewLine & "Нет - файл закроется", vbYesNoCancel, "Выбор дальнейших действий")
    Select Case result
        Case vbYes
    '       MsgBox ("Вы выбрали ДА")'открываю форму с вариантами шаблонов (ещё в работе)
        Case vbNo
    '       MsgBox ("Вы выбрали НЕТ")' Закрываю файл
            Workbooks("Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ_NEW.xlsx").Close SaveChanges:=False
        Case vbCancel
            MsgBox ("Вы выбрали [Отмена] - сейчас Вам предложат ввести номер коробки для определения в ней количества досье.") ' фильтрация с заполнением диапазона
    Range("AM2").Select
            
    KeyStr = InputBox("Введите ключевую последовательность символов")
              
    Set Begin = Range("AF2:AJ2")
    Set Rez = Range("AM2")
    
    TimeStart = Timer
    Rcount = CLng(Split(ActiveSheet.UsedRange.Address, "$")(4)) - Begin.Row + 1
    ArrFrom = Begin.Resize(Rcount, Begin.Count)
    ReDim ArrOut(1 To Rcount, 1 To 1)
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "\b" & KeyStr & "\b"
        
        For i = 1 To Rcount
            ArrOut(i, 1) = 0
            For j = 1 To Begin.Count
                If .Test(ArrFrom(i, j)) Then
                    ArrOut(i, 1) = 1
                    Exit For
                End If
            Next
        Next
    End With
    
    Rez.Resize(Rcount, 1) = ArrOut
'    MsgBox "Обработка данных завершена за " & Round(Timer - TimeStart, 2) & " сек."
            
'        Dim BoxNumber As Long
'        BoxNumber = InputBox("Пожалуйста, введите BoxNumber для обработки", "Запрос номера")
'        Debug.Print "BoxNumber= "; BoxNumber 'Ок
'
'        TimeStart = Timer ' засекаю время
'
'        Range("AM3").Select
'        Range("AM3:AM77777").FormulaR1C1 = "=Шапка_5.xlsm!CHK(RC[-7]:RC[-3] & BoxNumber)"
'        Debug.Print "AM3= "; Range("AM3").Value 'Error 2015
'        Range("AM3:AM77777").Select
'        Dim smallrng As Range ' далее замена формул на значения:
'        For Each smallrng In Selection.Areas
'            smallrng.Value = smallrng.Value
'        Next smallrng

        End Select
        
        '   Предлагаю фомулу вставить и заменить на значения. Лучше бы конечно на VBA....
        If MsgBox("Хотите запустить заполнение ячеек [AY4:AY77777] формулой ВПР с заменой на значения ?", vbYesNo, "Вопрос:") = vbNo Then Exit Sub
            
                  f1_Ожидайте.Show 0
                    DoEvents
                    Application.ScreenUpdating = False
                    Application.Wait Now + TimeSerial(0, 0, 0.6)
            
            Workbooks("Итог_ФКБ_Лазарев.xlsm").Activate   ' на передний план [Итог.xlsx]
            Worksheets("Лист1").Range("AY4:AY77777").FormulaLocal = "=ВПР(B4;'[Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ_NEW.xlsx]Лист1'!$B:$AM;38;0)"
            Application.ScreenUpdating = False
            Range("AY4:AY77777") = Range("AY4:AY77777").Value
            Application.ScreenUpdating = True
            With CreateObject("WScript.Shell")
            
                   Unload f1_Ожидайте
                     Application.ScreenUpdating = True
                
                .Run "mshta.exe vbscript:close(CreateObject(""WScript.shell"").Popup(""Формула в ячейки вставлена и заменена на значения"",5,""Информация:""))"
            End With
    ' Фильтр на 0 и удалить нули в столбце
              
    ' Сюда нужен ПрогрессБар нормальный
    
    Dim f2
    Const lSeconds7 As Long = 8
    MessageBoxTimeOut 0, "Обработка данных завершена за " & Round(Timer - TimeStart, 2) & " сек." & _
                  vbNewLine & " ", "Тайминг вычислений", _
                  vbInformation + vbOKOnly, 0&, lSeconds7 * 1000
End Sub
 
Function IsBookOpen(wbFullName As String) As Boolean
    Dim iFF As Integer, RetVal As Boolean
    iFF = FreeFile
    On Error Resume Next
    Open wbFullName For Random Access Read Write Lock Read Write As #iFF
    RetVal = (Err.Number <> 0)
    Close #iFF
    IsBookOpen = RetVal
End Function


