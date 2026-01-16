Attribute VB_Name = "m9_Автозаполнение_отчетов"
Option Explicit
Option Compare Text
Public КоличествоЗапросов As Long
Public КоличествоБумЗапросов As Long
Public lf As String
Public ДиапазонКлаймов  As Range

Sub СоздатьСписокКлаймов(control As IRibbonControl) '   Это первая часть Автозаполнения отчета по клаймам ПОИСК ОН
    
    Dim wBook As Workbook
    Dim sheet As Worksheet
'    Dim cell As Range
    Dim sName As String
    Dim Start As Date
    
    sName = "Клаймы" 'Создаю переменную, в которую помещаю имя листа.

    On Error Resume Next
        If Worksheets(sName) Is Nothing Then             'Действия, если листа нет
        If MsgBox("Внимание! " & _
                   vbNewLine & "Вы активировали программу запуска автозаполнения " & _
                   vbNewLine & "файла Отчёта по клаймам (Поиск ОН)." & _
                   vbNewLine & "Смени фильтр на СЕГОДНЯ! Продолжить?", vbYesNo) <> vbYes Then Exit Sub
            Worksheets.Add.Name = "Клаймы"
        End If
      On Error GoTo 0
                 Application.ScreenUpdating = False ' ОТКЛЮЧИЛ ЭКРАН
                  Worksheets("ппон").Select
    КоличествоЗапросов = Range("C5").Value
    КоличествоБумЗапросов = Range("C3").Value
                
    
    Sheets("4692").Select
    '   снимаю фильтры на активном листе:
        If ActiveSheet.FilterMode = True Then    'If wBook.ActiveSheet.FilterMode = True Then  не канает, обрезал код
             ActiveSheet.ShowAllData
        End If
    
'''    Dim ws As Worksheet       '    ТЕКУЩАЯ ДАТА В КРИТЕРИИ АВТОФИЛЬТРА !!!
'''    Dim lastRow As Long
'''    Dim filterColumn As Long
'''    Dim todayDate As Date
'''    Set ws = ActiveWorkbook.Sheets("4692")
'''
'''    filterColumn = 20 ' Укажите номер столбца для фильтрации (например, 1 для столбца A)
'''
'''    lastRow = ws.Cells(Rows.Count, filterColumn).End(xlUp).Row  ' Получаем последнюю строку с данными в столбце
'''
'''    todayDate = Date  ' Получаем текущую дату
'''
'''    ' Применяем автофильтр:
'''    With ws.Range(ws.Cells(1, filterColumn), ws.Cells(lastRow, filterColumn))
'''        .AutoFilter Field:=1, Criteria1:=todayDate
'''    End With

Dim СЕГОДНЯ As String
Dim todayDate As Date
todayDate = Date  ' Получаем текущую дату
СЕГОДНЯ = todayDate

                                                   '    СМЕНИ ДАТУ !!!
                                                   '    СМЕНИ ДАТУ !!!
                                                   '    СМЕНИ ДАТУ !!!
   Sheets("4692").Range("T1:T4693").AutoFilter Field:=20, Criteria1:=СЕГОДНЯ 'СТАВЛЮ фильтры на активном листе в 20-м столбце Т

 
    Columns("A:A").Select
    Selection.Copy
    Sheets("Клаймы").Select
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("4692").Select
    Columns("P:P").Select
    Application.CutCopyMode = False
    Selection.Copy
'               ActiveSheet.Range("$A$1:$T$4693").AutoFilter Field:=20 ' снимаю фильтры на активном листе в ДВАДЦАТОМ столбце Т
    Sheets("Клаймы").Select
    Columns("B:B").Select
            Start = Timer
                        Do While Timer < Start + 0.5 '0.5 = полсекунды
                            DoEvents
                        Loop
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    
    ActiveSheet.Range("$A$1:$B$135").AutoFilter Field:=2, Criteria1:="1"  ''поставил фильтр в столбце В на 1
             Start = Timer
                Do While Timer < Start + 0.5 '0.5 = полсекунды
                    DoEvents
                Loop
'If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
    ActiveWindow.ScrollRow = 1  'скролл наверх
    ' переместиться на первую видимую ячейку столбца:
        With Worksheets("Клаймы").AutoFilter.Range
       Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    End With
    
   ' Диапазон от активной ячейки до последней непустой ячейки внизу выделяется с помощью кода:
    Range(ActiveCell, ActiveCell.End(xlDown)).Select
    Selection.Copy
    
 On Error Resume Next
 sName = "Клаймы2"
        If Worksheets(sName) Is Nothing Then             'Действия, если листа нет
'            If MsgBox("На данном листе эта кнопка не работает," & _
                   vbNewLine & "создать нужный лист и заполнить данными?", vbYesNo) <> vbYes Then Exit Sub
            Worksheets.Add.Name = "Клаймы2"
        End If
      On Error GoTo 0
    Sheets("Клаймы2").Select
    Range("A1").Select
    ActiveSheet.Paste
    
'If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
    Sheets("Клаймы").Select
    ' Снимаю фильтры на активном листе во втором столбце В:
       ActiveSheet.Range("$A$1:$T$4693").AutoFilter Field:=2
      ActiveSheet.Range("$A$1:$B$135").AutoFilter Field:=2, Criteria1:="2"  ''поставил фильтр в столбце В на 2
    ' переместиться на первую видимую ячейку столбца:
       With Worksheets("Клаймы").AutoFilter.Range
           Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
            If Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Text = "" Then
               GoTo Инструкция1
            End If
        End With

    
   ' Диапазон от активной ячейки до последней непустой ячейки внизу выделяется с помощью кода:
    Range(ActiveCell, ActiveCell.End(xlDown)).Select
    Selection.Copy
    Sheets("Клаймы2").Select
    Dim iLastRow As Long
                iLastRow = Cells(Rows.Count, 1).End(xlUp).Row
                Cells(iLastRow + 1, 1).Select ' для выделения ячейки, находящейся в последней строке и 1 столбце, после последней заполненной ячейки
    ActiveSheet.Paste
'   вствка второй раз:
                iLastRow = Cells(Rows.Count, 1).End(xlUp).Row
                Cells(iLastRow + 1, 1).Select ' для выделения ячейки, находящейся в последней строке и 1 столбце, после последней заполненной ячейки
    ActiveSheet.Paste
    
Инструкция1:
    
    Sheets("Клаймы").Select  '  работаю по вставке ТРИ:
    ActiveSheet.Range("$A$1:$B$135").AutoFilter Field:=2, Criteria1:="3"  ''поставил фильтр в столбце В на 3
    
         With Worksheets("Клаймы").AutoFilter.Range
           Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
            If Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Text = "" Then
               GoTo Инструкция2
            End If
        End With


   ' Диапазон от активной ячейки до последней непустой ячейки внизу выделяется с помощью кода:
    Range(ActiveCell, ActiveCell.End(xlDown)).Select
    Selection.Copy
    Sheets("Клаймы2").Select
                iLastRow = Cells(Rows.Count, 1).End(xlUp).Row
                Cells(iLastRow + 1, 1).Select ' для выделения ячейки, находящейся в последней строке и 1 столбце, после последней заполненной ячейки
    ActiveSheet.Paste  'ОШИБКА ОШИБКА ОШИБКА ОШИБКА ОШИБКА ОШИБКА ОШИБКА ------ ЕСЛИ ПРИ КРИТЕРИИ 3 ВЫХОДИТ ЛИШЬ ОДИН КЛАЙМ ИЛИ НОЛЬ КЛАЙМОВ.
'   вствка второй раз:
                iLastRow = Cells(Rows.Count, 1).End(xlUp).Row
                Cells(iLastRow + 1, 1).Select ' для выделения ячейки, находящейся в последней строке и 1 столбце, после последней заполненной ячейки
    ActiveSheet.Paste
'   вствка третий раз:
                iLastRow = Cells(Rows.Count, 1).End(xlUp).Row
                Cells(iLastRow + 1, 1).Select ' для выделения ячейки, находящейся в последней строке и 1 столбце, после последней заполненной ячейки
    ActiveSheet.Paste
    
Инструкция2:         ' Метка "Инструкция2"
'If MsgBox("Все три фильтра отработаны, дальше?", vbYesNo) = vbNo Then Exit Sub

Sheets("4692").Select
    Sheets("4692").Range("T:T").AutoFilter Field:=20 ' снимаю фильтр на листе Sheets("4692") в ДВАДЦАТОМ столбце Т
'    Range("L1").Select
    
'    Application.DisplayAlerts = False 'Оключил предупреждения
''    Sheets("Клаймы").Delete           'удаляю Sheets(Клаймы)
'    Application.DisplayAlerts = True  'Включил предупреждения

Sheets("Клаймы2").Select
'   выделяю диапазон и сохраняю его в переменную:
    Range("A1").Select
'   Диапазон от активной ячейки до последней непустой ячейки внизу выделяется с помощью кода:
    Range(ActiveCell, ActiveCell.End(xlDown)).Select
    Selection.Copy
'    Range("B1").Select '    вставка диапазона в ячейку работает
'    ActiveSheet.Paste
'    Dim ДиапазонКлаймов As Range '    закомментил так как перенес объявление переменной в публичный доступ
    

     
     
     Set ДиапазонКлаймов = Selection 'Присваиваем перменной диапазон ячеек
'     If Not ДиапазонКлаймов Is Nothing Then
'    ' Выводим адрес диапазона в окно отладки
'    Debug.Print "Выделенный диапазон: " & ДиапазонКлаймов.Address
'
'    ' Выводим значения ячеек диапазона в окно отладки
'        Dim cell As Range
'            For Each cell In ДиапазонКлаймов
'                  Debug.Print "Ячейка " & cell.Address & ": " & cell.Value
'                Next cell
'            Else
'                Debug.Print "Диапазон не выделен."
'        End If
                      Application.ScreenUpdating = True ' ПОДКЛЮЧИЛ ЭКРАН
                      
'    If MsgBox("Запустить ЗаполнитьОтчетПоКлаймамОН?", vbYesNo) = vbNo Then Exit Sub

    
'    Sheets("Клаймы2").Delete '   удаляю Sheets(Клаймы2)
    Call ЗаполнитьОтчетПоКлаймамОН
End Sub


''Sub АвтозапускЗаполнитьОтчетФабрика(control As IRibbonControl)
''
''If MsgBox("Запустить старт по времени: БитриксЗавершение, ЗаполнитьОтчетФабрика, БитриксНачало?", vbYesNo) = vbNo Then Exit Sub
''
''        Application.OnTime TimeValue("17:02:00"), "БитриксЗавершение"
''
''        Application.OnTime TimeValue("05:42:00"), "ЗаполнитьОтчетФабрика"
''
''        Application.OnTime TimeValue("06:15:00"), "БитриксНачало"
''End Sub


Sub Авто_на_ночь(control As IRibbonControl)

If MsgBox("Запустить старт по времени: БитриксЗавершение, ЗаполнитьОтчетФабрика, БитриксНачало?", vbYesNo) = vbNo Then Exit Sub

'        Application.OnTime TimeValue("17:02:00"), "БитриксЗавершение"
'        Application.OnTime TimeValue("17:20:00"), "Добавить_B_текстовый_файл"
'
'        Application.OnTime TimeValue("05:42:00"), "ЗаполнитьОтчетФабрика"
        
        Application.OnTime TimeValue("06:18:00"), "БитриксНачало"
End Sub



Sub ЗаполнитьДинамикуФКБ(control As IRibbonControl)
    Dim Результат1 As String, Результат2 As String  'объявил переменные типа String
    Dim Кол_воКлаймов As String        'объявил локальную переменную
    Dim fl As Boolean                  'объявил локальную переменную
    Dim wb1 As Excel.Workbook, wb2 As Excel.Workbook        'объявил переменные для книг [Итог_ФКБ_Лазарев.xlsm] и [Динамика..xlsx]
    Dim UserName1 As String
'    Set wb1 = Workbooks.Open("C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\Итог_ФКБ_Лазарев.xlsm")        ' закомментишь- не работает при открытой книге(?)
'   Range("AY3").ClearContents        'не важно зачем
    Application.ScreenUpdating = False        'оператор VBA для отключения обновления экрана (=ускорение), он же фоновый режим(?))
    
'   подключаю функцию проверки IsBookOpen("wbFullName")на открытость/закрытость книги2 [Итог.xlsx]: ДОДЕЛАТЬ - ЕСЛИ ОТКРЫТ!!!
'   внимание! если в этом модуле нет описания этой функции и обращаешься к ней в другой модуль-пропиши Module5_ФункцияПроверки_Ореп_wb.IsBookOpen("C:\Users\.....
    fl = IsBookOpen("Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx")
    UserName1 = Environ("USERNAME")
    MsgBox "Сергей Александрович!" & vbNewLine & "Файл " & "Динамика 2025 Электрозаводская.xlsx'" & IIf(fl, " уже открыт", " НЕ ЗАНЯТ")
    
'   Кол_воКлаймов = Range("AX3")        'значение для переменной беру из листа [ппон] книги [Шапка_5.xlsm], или из:
    Кол_воКлаймов = InputBox("Количество отработанных клаймов:", "Заполнение Динамики", "Введите значение", 13000, 6000)  'присвоил переменной значение из InputBox
    If Кол_воКлаймов = "" Then Exit Sub
    If Кол_воКлаймов = "Введите значение" Then Exit Sub
    Set wb2 = Workbooks.Open("Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx")
    
    Dim sht, sht21, sht22         As Worksheet        'объявил переменную для листов
    Dim ИмяЛиста1, ИмяЛиста2  As String
    Dim currentmonthname As String      'объявил переменную для имени текущего месяца
    
    currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
    
    ИмяЛиста1 = (currentmonthname & " 2025")
    ИмяЛиста2 = ("учет " & currentmonthname & " 2025")
'    MsgBox "Имя Листа1 " & ИмяЛиста1 & vbNewLine & "Имя Листа2 " & ИмяЛиста2
    ' If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
    
    Set sht21 = wb2.Worksheets(ИмяЛиста1)        'установил ссылку на лист [ххххххх] [Динамика..xlsx]
    Set sht22 = wb2.Worksheets(ИмяЛиста2)         'установил ссылку на лист [учет хххххх 2025] [Динамика..xlsx]
    
    'РАБОТАЮ ПЕРВЫЙ ЛИСТ--------------------------------------------------------------------------------------------------------
    
    sht21.Activate        'активировал  лист [декабрь 2024] [Динамика...xlsx]
    Dim strAddress  As String        'объявил переменную для адреса ячейки
    Dim rng         As Range        'объявил переменную
    Dim dtToday     As Date        'объявил переменную для СЕГОДНЯ()
    dtToday = Date
    'начинаю искать ячейку по условию:
    Dim Cl As Range, Iskomoe$        'объявил переменную
    Iskomoe = "Лазарев С.А."
    Iskomoe = "*" & LCase(Iskomoe) & "*"
    
    For Each Cl In Range("B2:B130" & Range("B2").End(xlDown).Row)        ' ... Range("B2:B")... - почему-то не катит
        If LCase(Cl) Like Iskomoe And Cl.Offset(, -1) = dtToday Then Cl.Select        'выделил найденную ячейку
    Next
    
    If Not Intersect(ActiveCell, Range("B4:B130")) Is Nothing Then
        Cells(ActiveCell.Row, 30).Value = Кол_воКлаймов    'переход на 30 ячеек вправо и передача знач-я переменной в ячейку.
        Cells(ActiveCell.Row, 30).Select     'выделил активную ячейку
        strAddress = ActiveCell.Address        '   получаю адрес активной ячейки
        Результат1 = Range(ActiveCell.Address)        '    опреатор Set ! и все равно переменная обнуляется... :(
        Debug.Print "Aдрес актив яч-1 = "; (ActiveCell.Address)
        Debug.Print "Результ-1 = "; Результат1
    End If
    
'    sht11.Activate        'активировал  лист [ппон] книги [Шапка_5.xlsm]
'    Range("AY3").Value = Результат1        'сохраняю в ячейку листа [ппон] книги [Шапка_5.xlsm],
'    Range("AX3").ClearContents        'не важно
    sht21.Activate        'снова активировал  лист [декабрь 2024] [Динамика...xlsx]
    Set rng = Range("A1:AQ108").Find(what:=Результат1, LookIn:=xlValues, LookAt:=xlWhole)        'здесь What:=Результат — искомое значение, LookIn:=xlValues — поиск по значениям ячеек, LookAt:=xlWhole — полное совпадение.
        If Not rng Is Nothing Then
            '        MsgBox "Нашли!"
            rng.Select
        Else
            MsgBox "Не найдено."
        End If
    Range("AR19").Activate        ' отойду в сторонку)
    
    'РАБОТАЮ ВТОРОЙ ЛИСТ----------------------------------------------------------------------------------------------------------
    
    sht22.Activate        'активировал  лист [учет декабрь 2024] [Динамика...xlsx]
    
    For Each Cl In Range("B2:B108" & Range("B2").End(xlDown).Row)
        If LCase(Cl) Like Iskomoe And Cl.Offset(, -1) = dtToday Then Cl.Select
    Next
    
    If Not Intersect(ActiveCell, Range("B4:B108")) Is Nothing Then
        Cells(ActiveCell.Row, 8).Value = Кол_воКлаймов        'Целевая ячейка2
        Cells(ActiveCell.Row, 8).Select
        strAddress = ActiveCell.Address
        Результат2 = Range(ActiveCell.Address)
        Debug.Print "Aдрес актив яч-2 = "; (ActiveCell.Address)
        Debug.Print "Результ-2 = "; Результат2
    End If
    
    Set rng = Range("A1:AQ130").Find(what:=Результат2, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        '        MsgBox "Нашли!"
        rng.Select
    Else
        MsgBox "Не найдено."
    End If
    Range("A1").Activate        'отойду в сторонку,
'    Range("AE11:AQ37").Select   ' выделил диапазон
    Workbooks("Динамика 2025 Электрозаводская.xlsx").Save        ' сохраняю изменения в книге
    Set wb2 = Nothing        'удаляю присвоенное значение переменной Workbook
    Workbooks("Динамика 2025 Электрозаводская.xlsx").Close SaveChanges:=False        'закрываю без сохранения wb [Динамика...xlsx]
    Application.ScreenUpdating = True        'включаю обновление экрана.
    
     MsgBox "The value of the target cell in the worksheet [май 2025] = " & Результат1 & _
        vbNewLine & "Значение целевой ячейки на листе [учет май 2025] = " & Результат2, vbOKOnly, _
        "Проверка заполнения ячеек на двух листах файла [Динамика...]"
    
End Sub



Sub ЗаполнитьДинамикуГПБ(control As IRibbonControl)
    Dim Результат1 As String, Результат2 As String  'объявил переменные типа String
    Dim Кол_воКлаймов As String        'объявил локальную переменную
    Dim fl As Boolean                  'объявил локальную переменную
    Dim wb1 As Excel.Workbook, wb2 As Excel.Workbook        'объявил переменные для книг [Итог_ФКБ_Лазарев.xlsm] и [Динамика..xlsx]
    Dim UserName1 As String
    
    If MsgBox("Заполнение Динамики ГПБ еще в разработке. Продолжить?", vbYesNo, "Имя приложения: " & Application.Name) = vbNo Then Exit Sub
'    Set wb1 = Workbooks.Open("C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\Итог_ФКБ_Лазарев.xlsm")        ' закомментишь- не работает при открытой книге(?)
'   Range("AY3").ClearContents        'не важно зачем
    Application.ScreenUpdating = False        'оператор VBA для отключения обновления экрана (=ускорение), он же фоновый режим(?))
    
'   подключаю функцию проверки IsBookOpen("wbFullName")на открытость/закрытость книги2 [Итог.xlsx]: ДОДЕЛАТЬ - ЕСЛИ ОТКРЫТ!!!
'   внимание! если в этом модуле нет описания этой функции и обращаешься к ней в другой модуль-пропиши Module5_ФункцияПроверки_Ореп_wb.IsBookOpen("C:\Users\.....
    fl = IsBookOpen("Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx")
    UserName1 = Environ("USERNAME")
    MsgBox "Сергей Александрович!" & vbNewLine & "Файл " & "Динамика 2025 Электрозаводская.xlsx'" & IIf(fl, " уже открыт", " НЕ ЗАНЯТ")
    
'   Кол_воКлаймов = Range("AX3")        'значение для переменной беру из листа [ппон] книги [Шапка_5.xlsm], или из:
    Кол_воКлаймов = InputBox("Количество отработанных клаймов:", "Заполнение Динамики", "Введите значение", 13000, 6000)  'присвоил переменной значение из InputBox
    If Кол_воКлаймов = "" Then Exit Sub
    If Кол_воКлаймов = "Введите значение" Then Exit Sub
    Set wb2 = Workbooks.Open("Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx")
    
    Dim sht, sht21, sht22         As Worksheet        'объявил переменную для листов
    Dim ИмяЛиста1, ИмяЛиста2  As String
    Dim currentmonthname As String      'объявил переменную для имени текущего месяца
    
    currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
    
    ИмяЛиста1 = (currentmonthname & " 2025")
    ИмяЛиста2 = ("учет " & currentmonthname & " 2025")
'    MsgBox "Имя Листа1 " & ИмяЛиста1 & vbNewLine & "Имя Листа2 " & ИмяЛиста2
    ' If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
    
    Set sht21 = wb2.Worksheets(ИмяЛиста1)        'установил ссылку на лист [ххххххх] [Динамика..xlsx]
    Set sht22 = wb2.Worksheets(ИмяЛиста2)         'установил ссылку на лист [учет хххххх 2025] [Динамика..xlsx]
    
    'РАБОТАЮ ПЕРВЫЙ ЛИСТ--------------------------------------------------------------------------------------------------------
    
    sht21.Activate        'активировал  лист [декабрь 2024] [Динамика...xlsx]
    Dim strAddress  As String        'объявил переменную для адреса ячейки
    Dim rng         As Range        'объявил переменную
    Dim dtToday     As Date        'объявил переменную для СЕГОДНЯ()
    dtToday = Date
    'начинаю искать ячейку по условию:
    Dim Cl As Range, Iskomoe$        'объявил переменную
    Iskomoe = "Лазарев С.А."
    Iskomoe = "*" & LCase(Iskomoe) & "*"
    
    For Each Cl In Range("B2:B130" & Range("B2").End(xlDown).Row)        ' ... Range("B2:B")... - почему-то не катит
        If LCase(Cl) Like Iskomoe And Cl.Offset(, -1) = dtToday Then Cl.Select        'выделил найденную ячейку
    Next
    
    If Not Intersect(ActiveCell, Range("B4:B130")) Is Nothing Then
        Cells(ActiveCell.Row, 30).Value = Кол_воКлаймов    'переход на 30 ячеек вправо и передача знач-я переменной в ячейку.
        Cells(ActiveCell.Row, 30).Select     'выделил активную ячейку
        strAddress = ActiveCell.Address        '   получаю адрес активной ячейки
        Результат1 = Range(ActiveCell.Address)        '    опреатор Set ! и все равно переменная обнуляется... :(
        Debug.Print "Aдрес актив яч-1 = "; (ActiveCell.Address)
        Debug.Print "Результ-1 = "; Результат1
    End If
    
'    sht11.Activate        'активировал  лист [ппон] книги [Шапка_5.xlsm]
'    Range("AY3").Value = Результат1        'сохраняю в ячейку листа [ппон] книги [Шапка_5.xlsm],
'    Range("AX3").ClearContents        'не важно
    sht21.Activate        'снова активировал  лист [декабрь 2024] [Динамика...xlsx]
    Set rng = Range("A1:AQ108").Find(what:=Результат1, LookIn:=xlValues, LookAt:=xlWhole)        'здесь What:=Результат — искомое значение, LookIn:=xlValues — поиск по значениям ячеек, LookAt:=xlWhole — полное совпадение.
        If Not rng Is Nothing Then
            '        MsgBox "Нашли!"
            rng.Select
        Else
            MsgBox "Не найдено."
        End If
    Range("AR19").Activate        ' отойду в сторонку)
    
    'РАБОТАЮ ВТОРОЙ ЛИСТ----------------------------------------------------------------------------------------------------------
    
    sht22.Activate        'активировал  лист [учет декабрь 2024] [Динамика...xlsx]
    
    For Each Cl In Range("B2:B108" & Range("B2").End(xlDown).Row)
        If LCase(Cl) Like Iskomoe And Cl.Offset(, -1) = dtToday Then Cl.Select
    Next
    
    If Not Intersect(ActiveCell, Range("B4:B108")) Is Nothing Then
        Cells(ActiveCell.Row, 8).Value = Кол_воКлаймов        'Целевая ячейка2
        Cells(ActiveCell.Row, 8).Select
        strAddress = ActiveCell.Address
        Результат2 = Range(ActiveCell.Address)
        Debug.Print "Aдрес актив яч-2 = "; (ActiveCell.Address)
        Debug.Print "Результ-2 = "; Результат2
    End If
    
    Set rng = Range("A1:AQ130").Find(what:=Результат2, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        '        MsgBox "Нашли!"
        rng.Select
    Else
        MsgBox "Не найдено."
    End If
    Range("A1").Activate        'отойду в сторонку,
'    Range("AE11:AQ37").Select   ' выделил диапазон
    Workbooks("Динамика 2025 Электрозаводская.xlsx").Save        ' сохраняю изменения в книге
    Set wb2 = Nothing        'удаляю присвоенное значение переменной Workbook
    Workbooks("Динамика 2025 Электрозаводская.xlsx").Close SaveChanges:=False        'закрываю без сохранения wb [Динамика...xlsx]
    Application.ScreenUpdating = True        'включаю обновление экрана.
    
     MsgBox "The value of the target cell in the worksheet [май 2025] = " & Результат1 & _
        vbNewLine & "Значение целевой ячейки на листе [учет май 2025] = " & Результат2, vbOKOnly, _
        "Проверка заполнения ячеек на двух листах файла [Динамика...]"
    
End Sub




Public Sub ЗаполнитьДинамикуОН(control As IRibbonControl)
 MsgBox "Не найдено."
  Dim Результат1 As String, Результат2 As String  'объявил переменные типа String
    Dim Кол_воКлаймов As String        'объявил локальную переменную
    Dim fl As Boolean                  'объявил локальную переменную
    Dim wb1 As Excel.Workbook, wb2 As Excel.Workbook        'объявил переменные для книг [Итог_ФКБ_Лазарев.xlsm] и [Динамика..xlsx]
    Dim UserName1 As String
'    Set wb1 = Workbooks.Open("C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\Итог_ФКБ_Лазарев.xlsm")        ' закомментишь- не работает при открытой книге(?)
'   Range("AY3").ClearContents        'не важно зачем
    Application.ScreenUpdating = False        'оператор VBA для отключения обновления экрана (=ускорение), он же фоновый режим(?))
    
'   подключаю функцию проверки IsBookOpen("wbFullName")на открытость/закрытость книги2 [Итог.xlsx]: ДОДЕЛАТЬ - ЕСЛИ ОТКРЫТ!!!
'   внимание! если в этом модуле нет описания этой функции и обращаешься к ней в другой модуль-пропиши Module5_ФункцияПроверки_Ореп_wb.IsBookOpen("C:\Users\.....
    fl = IsBookOpen("Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx")
    UserName1 = Environ("USERNAME")
    MsgBox "Сергей Александрович!" & vbNewLine & "Файл " & "Динамика 2025 Электрозаводская.xlsx'" & IIf(fl, " уже открыт", " НЕ ЗАНЯТ")
    
'   Кол_воКлаймов = Range("AX3")        'значение для переменной беру из листа [ппон] книги [Шапка_5.xlsm], или из:
    Кол_воКлаймов = InputBox("Количество отработанных клаймов:", "Заполнение Динамики", "Введите значение", 13000, 6000)  'присвоил переменной значение из InputBox
    If Кол_воКлаймов = "" Then Exit Sub
    If Кол_воКлаймов = "Введите значение" Then Exit Sub
    Set wb2 = Workbooks.Open("Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx")
    
    Dim sht, sht21, sht22         As Worksheet        'объявил переменную для листов
    Dim ИмяЛиста1, ИмяЛиста2  As String
    Dim currentmonthname As String      'объявил переменную для имени текущего месяца
    
    currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
    
    ИмяЛиста1 = (currentmonthname & " 2025")
    ИмяЛиста2 = ("учет " & currentmonthname & " 2025")
'    MsgBox "Имя Листа1 " & ИмяЛиста1 & vbNewLine & "Имя Листа2 " & ИмяЛиста2
    ' If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
    
    Set sht21 = wb2.Worksheets(ИмяЛиста1)        'установил ссылку на лист [ххххххх] [Динамика..xlsx]
    Set sht22 = wb2.Worksheets(ИмяЛиста2)         'установил ссылку на лист [учет хххххх 2025] [Динамика..xlsx]
    
    'РАБОТАЮ ПЕРВЫЙ ЛИСТ--------------------------------------------------------------------------------------------------------
     sht21.Activate        'активировал  лист [декабрь 2024] [Динамика...xlsx]
    Dim strAddress  As String        'объявил переменную для адреса ячейки
    Dim rng         As Range        'объявил переменную
    Dim dtToday     As Date        'объявил переменную для СЕГОДНЯ()
    dtToday = Date
    'начинаю искать ячейку по условию:
    Dim Cl As Range, Iskomoe$        'объявил переменную
    Iskomoe = "Лазарев С.А."
    Iskomoe = "*" & LCase(Iskomoe) & "*"
    
    For Each Cl In Range("B2:B130" & Range("B2").End(xlDown).Row)        ' ... Range("B2:B")... - почему-то не катит
        If LCase(Cl) Like Iskomoe And Cl.Offset(, -1) = dtToday Then Cl.Select        'выделил найденную ячейку
    Next
    
    If Not Intersect(ActiveCell, Range("B4:B130")) Is Nothing Then
        Cells(ActiveCell.Row, 30).Value = Кол_воКлаймов    'переход на 30 ячеек вправо и передача знач-я переменной в ячейку.
        Cells(ActiveCell.Row, 30).Select     'выделил активную ячейку
        strAddress = ActiveCell.Address        '   получаю адрес активной ячейки
        Результат1 = Range(ActiveCell.Address)        '    опреатор Set ! и все равно переменная обнуляется... :(
        Debug.Print "Aдрес актив яч-1 = "; (ActiveCell.Address)
        Debug.Print "Результ-1 = "; Результат1
    End If
    
'    sht11.Activate        'активировал  лист [ппон] книги [Шапка_5.xlsm]
'    Range("AY3").Value = Результат1        'сохраняю в ячейку листа [ппон] книги [Шапка_5.xlsm],
'    Range("AX3").ClearContents        'не важно
    sht21.Activate        'снова активировал  лист [декабрь 2024] [Динамика...xlsx]
    Set rng = Range("A1:AQ108").Find(what:=Результат1, LookIn:=xlValues, LookAt:=xlWhole)        'здесь What:=Результат — искомое значение, LookIn:=xlValues — поиск по значениям ячеек, LookAt:=xlWhole — полное совпадение.
        If Not rng Is Nothing Then
            '        MsgBox "Нашли!"
            rng.Select
        Else
            MsgBox "Не найдено."
        End If
    Range("AR19").Activate        ' отойду в сторонку)
    
    'РАБОТАЮ ВТОРОЙ ЛИСТ----------------------------------------------------------------------------------------------------------
    
    sht22.Activate        'активировал  лист [учет декабрь 2024] [Динамика...xlsx]
    
    For Each Cl In Range("B2:B108" & Range("B2").End(xlDown).Row)
        If LCase(Cl) Like Iskomoe And Cl.Offset(, -1) = dtToday Then Cl.Select
    Next
    
    If Not Intersect(ActiveCell, Range("B4:B108")) Is Nothing Then
        Cells(ActiveCell.Row, 8).Value = Кол_воКлаймов        'Целевая ячейка2
        Cells(ActiveCell.Row, 8).Select
        strAddress = ActiveCell.Address
        Результат2 = Range(ActiveCell.Address)
        Debug.Print "Aдрес актив яч-2 = "; (ActiveCell.Address)
        Debug.Print "Результ-2 = "; Результат2
    End If
    
    Set rng = Range("A1:AQ130").Find(what:=Результат2, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        '        MsgBox "Нашли!"
        rng.Select
    Else
        MsgBox "Не найдено."
    End If
    Range("A1").Activate        'отойду в сторонку,
'    Range("AE11:AQ37").Select   ' выделил диапазон
    Workbooks("Динамика 2025 Электрозаводская.xlsx").Save        ' сохраняю изменения в книге
    Set wb2 = Nothing        'удаляю присвоенное значение переменной Workbook
    Workbooks("Динамика 2025 Электрозаводская.xlsx").Close SaveChanges:=False        'закрываю без сохранения wb [Динамика...xlsx]
    Application.ScreenUpdating = True        'включаю обновление экрана.
    
     MsgBox "The value of the target cell in the worksheet [май 2025] = " & Результат1 & _
        vbNewLine & "Значение целевой ячейки на листе [учет май 2025] = " & Результат2, vbOKOnly, _
        "Проверка заполнения ячеек на двух листах файла [Динамика...]"
End Sub


'=========================== начало Sub ЗаполнитьОтчетПоСканирФКБ  ОТЧЕТ по КЛАЙМАМ\========================
'=========================== начало Sub ЗаполнитьОтчетПоСканирФКБ  ОТЧЕТ по КЛАЙМАМ\========================

'Option Explicit
'Option Compare Text
'Public PapkaSkanov, lf As String ' перменная становится доступной для использования (видимой) в любом модуле и любой процедуре проекта
'Public КоличСканов As Variant    ' не обязательно Public

Public Sub ЗаполнитьОтчетПоСканированиюФКБ(control As IRibbonControl) '   control As IRibbonControl
'Sub ЗаполнитьОтчетПоСканированиюФКБ() '
'    Dim КоличСканов As Variant
    Dim Segodnja As String
    Dim SegodnjaADRES As String
    Dim FSO As Object, fFolders As Object, fFolder As Object
    Dim sFolderName As String
    Dim FolderPath As String
    Dim fl As String
    Dim wBook As Workbook
    Dim КоличСканов As Variant
    
' If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
'--------------------------------------------


    Dim oFD As FileDialog '...............ВЫБОР ПАПКИ Sub ВыборПапкиСПапкамиСканов(control As IRibbonControl)
    On Error Resume Next
    'назначаем переменной ссылку на экземпляр диалога
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    With oFD 'используем короткое обращение к объекту
    'так же можно без oFD
    'With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Выбрать папку с отчетами" '"заголовок окна диалога
        .ButtonName = "Выбрать папку"
        .Filters.Clear 'очищаем установленные ранее типы файлов
        .InitialFileName = "C:\Users\s.lazarev\Desktop\" '"назначаем первую папку отображения
        .InitialView = msoFileDialogViewLargeIcons 'вид диалогового окна(доступно 9 вариантов)
        If .Show = 0 Then Exit Sub 'показывает диалог
        'цикл по коллекции выбранных в диалоге файлов
        PapkaSkanov = .SelectedItems(1) 'считываем путь к папке
'        MsgBox "Выбрана папка: '" & PapkaSkanov & "'", vbInformation, "Сообщение"
'        MsgBox "Выбрана папка " & PapkaSkanov
    End With     '...............ВЫБОР ПАПКИ конец
    
On Error GoTo 0

    On Error GoTo ErrHandler
'   Application.DisplayAlerts = False
'   Application.ScreenUpdating чуть ниже
   
    FolderPath = PapkaSkanov
    ' получаем характеристики папки = PapkaSkanov
    Set FSO = CreateObject("Scripting.FileSystemObject")
    КоличСканов = FSO.GetFolder(FolderPath).Subfolders.Count
         MsgBox "Выбрана папка " & PapkaSkanov & vbNewLine & vbNewLine & "Как бэ сегодня папок со сканами = " & КоличСканов
         
    Dim pathОтчета As String
    Dim pathОтчета_с_currentmonthname As String
    Dim currentmonthname As String
    
    currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
    pathОтчета = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\Отчет по клаймам за июнь 2025.xlsx" ' полное имя Отчета по клаймам
    pathОтчета_с_currentmonthname = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\" & "Отчет по клаймам за " & currentmonthname & " 2025.xlsx"
    
    '  проверка, открыт ли наш файл (без специальной функции Function IsBookOpen(wbFullName As String) As Boolean)
    Dim sWBName As String
    Dim bChk As Boolean
    
    sWBName = pathОтчета_с_currentmonthname
    
    For Each wBook In Workbooks
        If wBook.Name = sWBName Then
            Set wBook = Workbooks("Отчет по клаймам за " & currentmonthname & " 2025.xlsx")
            wBook.Windows(1).Activate 'вытягиваем на первый план
        End If
    Next
    
    If bChk = False Then
        Set wBook = Workbooks.Open(FileName:=pathОтчета_с_currentmonthname)  ' открыл
    End If
 '  конец кода проверки, открыт ли наш файл
    
    With Application  ' размещаю окно  в координатах:
        .WindowState = xlNormal
        .Width = 1420 ' ШИРИНА окна
        .Height = 445 ' ВЫСОТА окна
        .Left = 0     ' ЛЕВАЯ точка
        .Top = 0    ' ВЕРХНЯЯ точка
    End With
    Worksheets("отчет за день").Select
     
'                         Application.ScreenUpdating = False  ' можно отрубить отображение процессов на экране

'   снимаю фильтры на активном листе:
        If wBook.ActiveSheet.FilterMode = True Then
             wBook.ActiveSheet.ShowAllData
        End If
'   Загрузка списка папок в указанный диапазон: Function SubFoldersCollection(ByVal folderPath As String, Optional ByVal Mask As String = "*") As Collection
     Dim i           As Long
'    On Error GoTo ErrHandler
 
    Dim L           As String
    L = PapkaSkanov
 
    Dim coll        As Collection
    Set coll = SubFoldersCollection(L)
 
       With ActiveWorkbook.Worksheets("отчет за день")    ' замените на конкретное имя вашего листа
    
           Dim nextRow As Long
           nextRow = .Cells(.Rows.Count, "D").End(xlUp).Row + 1
           If .Cells(1, "D").Value = "" Then nextRow = 1
    
           For i = 1 To coll.Count
               .Cells(nextRow, 4).Value = coll(i)
               nextRow = nextRow + 1
           Next i
    
       End With
    Application.Wait (Now + TimeValue("0:00:01"))  ' Пауза
'If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
    Dim iLastRow As Long
                iLastRow = Cells(Rows.Count, 2).End(xlUp).Row
                Cells(iLastRow + 1, 2).Select ' для выделения ячейки, находящейся в последней строке и втором столбце, после последней заполненной ячейки
            
            With Selection.Validation '   в выбранной ячейке создать выпадающий список и выбрать нужное значение для неё
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Ващенко Оксана Васильевна,Лазарев Сергей Александрович"
            End With
            
    Application.Wait (Now + TimeValue("0:00:01")) ' Пауза
    
    Dim username As String
    username = Environ("UserName")  ' Получаем имя пользователя компьютера.
        
    Select Case Environ("UserName")
        Case "v.petrov"
              Cells(iLastRow + 1, 2) = "Ващенко Оксана Васильевна"
        Case "s.lazarev"
            Cells(iLastRow + 1, 2) = "Лазарев Сергей Александрович"
        Case "g.sidorov"
              Cells(iLastRow + 1, 2) = "Гоша Сидоров"
        End Select
        
      If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
        
        Cells(iLastRow + 1, 3).Select ' — выбор ячейки, расположенной в строке, следующей за последней, и в 3-м столбце.
        
        With Selection.Validation '   в выбранной ячейке создать выпадающий список и выбрать нужное значение для неё
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=тариф!$A$2:$A$10"
        End With
              Cells(iLastRow + 1, 3) = "Изъятие документов из поступивших от контрагента досье, согласно матрицы компании, их сканирование и переименование  ПРОСТОЕ ДОСЬЕ"
              Cells(iLastRow + 1, 6).Value = Date ' СЕГОДНЯ
                
'     If MsgBox("Протянуть 'Лазарев Сергей Александрович' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
     Cells(ActiveCell.Row, 2).Select 'перейти во второй столбец этой строки
'     If MsgBox("Протянуть 'Лазарев Сергей Александрович' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
     ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillDefault

'     If MsgBox("Протянуть 'Изъятие документов из....' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
     Cells(ActiveCell.Row, 3).Select 'перейти во третий столбец этой строки
     ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillDefault

     Cells(ActiveCell.Row, 6).Select 'перейти в шестой столбец этой строки
'     If MsgBox("Протянуть сегодня вниз на (КоличСканов-1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
     ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillCopy

     Cells(ActiveCell.Row, 4).Select 'перейти в четвертый  столбец этой строки
     Range(Selection, Selection.End(xlDown)).Select   'Выделяю всё вниз от этих ячеек включительно
   
    Application.ScreenUpdating = True

'   Подсчитать значения в текущем выделенном диапазоне
    Dim rowCount As Long
    rowCount = Selection.Count
    MsgBox "Количество значений в выбранном диапазоне: " & rowCount
    
    If MsgBox("Сохранить файл с внесёнными данными?", vbYesNo) <> vbYes Then GoTo Skip
    'сохранить текущую книгу
    ActiveWorkbook.Save
Skip:
    'Закрыть книгу
    If MsgBox("Закрыть ""Отчет по клаймам""?", vbYesNo) <> vbYes Then Exit Sub
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
ErrHandler:
    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Ошибка выполнения"
   On Error GoTo 0
   Exit Sub
 End Sub
 

 
Sub МакросПоискаТекстаВСтолбце() ' макро поиска в столбце выражения. Строку из этого макро использовал в проге выше
Dim r1 As Range, r2 As Range
Worksheets("отчет за день").Select
    Set r1 = Columns("B:B").Find("Сверка документов от контрагентов по кол-ву должников")
    If Not r1 Is Nothing Then r1.Copy
    
End Sub

Function NextFold(p As Variant, ByRef flCount As Long, ByRef fldCount As Long)
    If Not (p Is Nothing) Then
        For Each iFolderItem In p.Items
            If Not iFolderItem.IsFolder Then
                flCount = flCount + 1
            Else
                fldCount = fldCount + 1
                Call NextFold(CreateObject("Shell.Application").Namespace(CVar(iFolderItem.Path)), flCount, fldCount)
            End If
        Next iFolderItem
    End If
End Function

' Для Sub ЗаполнитьОтчетПоСканированию()а именно подпроги загрузки списка папок от Николая
Function SubFoldersCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "*") As Collection
    Dim curfold     As Object
    Dim Folder      As Object
    On Error GoTo ErrHandler
 
    Set SubFoldersCollection = New Collection
 
    Dim FSO         As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
 
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    Set curfold = FSO.GetFolder(FolderPath)
 
    If Not curfold Is Nothing Then
 
        For Each Folder In curfold.Subfolders
 
            If Folder.Path Like FolderPath & Mask Then
 
                ' только имена подпапок
                SubFoldersCollection.Add Folder.Name
 
                '                ' полные пути к подпапкам с именами подпапок
                '                SubFoldersCollection.Add folder.Path & "\"
            End If
 
        Next Folder
 
    End If
 
    Set FSO = Nothing
    Set curfold = Nothing
    Exit Function
 
ErrHandler:
    MsgBox "Ошибка при получении подпапок: " & Err.Description, vbCritical, "Ошибка"
    Set SubFoldersCollection = Nothing
    Set FSO = Nothing
    Set curfold = Nothing
End Function


Sub ЗаполнитьСейчасОтчетФабрика(control As IRibbonControl)
    Call ЗаполнитьОтчетФабрика
End Sub


Public Sub ЗаполнитьОтчетФабрика()
   
    Dim КоличСканов As Variant
    Dim Segodnja As String
    Dim SegodnjaADRES As String
    Dim FSO As Object, fFolders As Object, fFolder As Object
    Dim sFolderName As String
    Dim FolderPath As String
    Dim fl As String
    Dim wBook As Workbook
 
    
' If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
'--------------------------------------------


    On Error GoTo ErrHandler
    КоличСканов = Range("C2")
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

      
    Dim pathОтчета As String
    Dim pathОтчета_с_currentmonthname As String
    Dim currentmonthname As String
    
    currentmonthname = Format(Date, "mmmm") ' получаем полное название текущего месяца в формате "Июль"
    pathОтчета = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\Для руководителей\Операционная отчетность фабрики\2025\Отчет по фабрикe.New (Июнь 2025).xlsx" ' полное имя Отчета по клаймам
    pathОтчета_с_currentmonthname = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\Для руководителей\Операционная отчетность фабрики\2025\" & "Отчет по фабрикe.New (" & currentmonthname & " 2025).xlsx"
    

   '    подключаю функцию проверки IsBookOpen("wbFullName")на открытость/закрытость книги:
Проверка:
    fl = IsBookOpen(pathОтчета_с_currentmonthname)

'    MsgBox "Внимание! " & vbNewLine & "файл " & IIf(fl, " уже открыт", " не занят")

    If fl Then ' если уже кем-то открыт
'        MsgBox "Внимание! " & vbNewLine & "файл уже открыт, ждем 20 сек и снова проверяем"
        Application.Wait Now + TimeValue("00:00:20")
        GoTo Проверка
    Else
    Set wBook = Workbooks.Open(FileName:="Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\Для руководителей\Операционная отчетность фабрики\2025\" & "Отчет по фабрикe.New (" & currentmonthname & " 2025).xlsx")  ' открыл
    End If  '  конец кода проверки, открыт ли наш файл
    
 
 Application.ScreenUpdating = False    ' отключить отображение процессов на экране
 
    With Application  ' размещаю окно  в координатах:
        .WindowState = xlNormal
        .Width = 1420 ' ШИРИНА окна
        .Height = 445 ' ВЫСОТА окна
        .Left = 0     ' ЛЕВАЯ точка
        .Top = 0    ' ВЕРХНЯЯ точка
    End With
    Worksheets("Детализация").Select


'   снимаю фильтры на активном листе:
        If wBook.ActiveSheet.FilterMode = True Then
             wBook.ActiveSheet.ShowAllData
        End If
        
     
    Dim iLastRow As Long
                iLastRow = Cells(Rows.Count, 1).End(xlUp).Row
                Cells(iLastRow + 1, 1).Select ' для выделения ячейки, находящейся в последней строке и первом столбце, после последней заполненной ячейки
            
            With Selection.Validation '   в выбранной ячейке создать выпадающий список и выбрать нужное для неё значение.
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=Справочник!$J$2:$J$6"

            End With
            
    Application.Wait (Now + TimeValue("0:00:01")) ' Пауза
    Cells(iLastRow + 1, 1) = "Отдел архива длительного хранения"

    
            Dim currentTime As Date ' ЗАПОЛНИТЬ ЯЧЕЙКУ СТОЛБЦА С ДАТОЙ. УСЛОВИЕ: ЕСЛИ текущее время находится в промежутке с 7:00 до 17:00
            Dim startTime As Date   ' ТО ЗАПОЛНЯЕМ СЕГОДНЯШНЮЮ ДАТУ, ФОРМУЛА СЕГОДНЯ()
            Dim endTime As Date     ' ЕСЛИ НЕТ, ТО ВЧЕРАШНЮЮ ДАТУ: .Value = Date - 1 (Для ночного заполнения Отчета по фабрике)
            
            currentTime = Time ' Получаем текущее время
            
            ' Устанавливаем время начала (7:00)
            startTime = TimeValue("07:00:00")
            
            ' Устанавливаем время окончания (17:00)
            endTime = TimeValue("17:00:00")
            
            ' Проверяем, находится ли текущее время в заданном промежутке
                If currentTime >= startTime And currentTime <= endTime Then
        '        MsgBox "Текущее время находится в промежутке с 7:00 до 17:00"
                Cells(iLastRow + 1, 2).Value = Date
                    Else
        '            MsgBox "Текущее время не находится в промежутке с 7:00 до 17:00"
                    Cells(iLastRow + 1, 2).Value = Date - 1
                End If
    

    Dim username As String
    username = Environ("UserName")  ' Получаем имя пользователя компьютера.
        
    Select Case Environ("UserName")
        Case "v.petrov"
              Cells(iLastRow + 1, 3) = "Ващенко Оксана Васильевна"
        Case "s.lazarev"
            Cells(iLastRow + 1, 3) = "Лазарев Сергей Александрович"
        Case "g.sidorov"
              Cells(iLastRow + 1, 3) = "Гоша Сидоров"
        End Select
        
     
     Cells(iLastRow + 1, 4).Select ' для выделения ячейки, находящейся в последней строке и первором столбце, после последней заполненной ячейки
            
            With Selection.Validation '   в выбранной ячейке создать выпадающий список и выбрать нужное значение для неё
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=Справочник!$B$2:$B$36"
            End With
        Cells(iLastRow + 1, 4) = "Изъятие документов из поступивших от контрагента досье, согласно матрицы компании, их сканирование и переименование  ПРОСТОЕ ДОСЬЕ "
                   
'        If MsgBox("Задам вопрос: Дальше?", vbYesNo) = vbNo Then Exit Sub
       
              Cells(iLastRow + 1, 5).Value = КоличСканов
        
'         Worksheets("Детализация").Select
    ActiveWindow.ScrollRow = 1
    Range("B1").Select
    
'    If MsgBox("Сохранить файл с внесёнными данными?", vbYesNo) <> vbYes Then GoTo Skip

    ActiveWorkbook.Save
    
   Dim strFile_Path As String
       username = Environ("UserName")
       strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
   Open strFile_Path For Append As #1
   Print #1, " "
   Print #1, "Отчёт по фабрике заполнил " & username & " " & Now   '
   Close #1
    
    Application.ScreenUpdating = True
    
    
    Dim EmailApp As Outlook.Application
                    Dim Source As String
                    Set EmailApp = New Outlook.Application
                    Dim EmailItem As Outlook.MailItem
                    Set EmailItem = EmailApp.CreateItem(olMailItem)
                    EmailItem.To = "s.lazarev@bsv.legal"
                    EmailItem.cc = " "
                    EmailItem.BCC = " "
                    EmailItem.Subject = "Отчёт по фабрике"
                    EmailItem.HTMLBody = "Добрый день," & "<br>" & "Отчёт по фабрике заполнен " & Now & "<br>" & "<br>" & "<br>" & "<br>" & "С уважением," & "<br>" & "Лазарев Сергей Александрович" & "<br>" & "Специалист" & "<br>" & "Отдела архива длительного хранения" & "<br>" & "ООО ПКО ""БСВ""" & "<br>" & "8-952-900-99-49" & "<br>" & "s.lazarev@bsv.legal"
                '    Source = "" 'ThisWorkbook.FullName
                '    EmailItem.Attachments.Add Source
                    EmailItem.send
    
Skip:
    'Закрыть книгу
'    If MsgBox("Закрыть ""Отчет по клаймам""?", vbYesNo) <> vbYes Then Exit Sub
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
ErrHandler:
    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Ошибка выполнения"
   On Error GoTo 0
   Exit Sub
 End Sub

'Вариант 3: По просьбам читателей -код, который проверяет открыта ли книга независимо от её месторасположения
'и используемого приложения Excel. Книга может быть открыта другим пользователем (если книга на сервере),
'в другом экземпляре Excel или в этом же экземпляре Excel.
'
Function IsBookOpen(wbFullName As String) As Boolean
    Dim iFF As Integer, RetVal As Boolean
    iFF = FreeFile
    On Error Resume Next
    Open wbFullName For Random Access Read Write Lock Read Write As #iFF
    RetVal = (Err.Number <> 0) ' вот это что. если ошибки нет, то 0
    Close #iFF
    IsBookOpen = RetVal
End Function

