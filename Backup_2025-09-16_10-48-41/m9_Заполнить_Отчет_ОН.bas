Attribute VB_Name = "m9_Заполнить_Отчет_ОН"
Option Explicit
Option Compare Text

'Public Sub ЗаполнитьОтчетПоКлаймамОН(control As IRibbonControl) '   control As IRibbonControl
Public Sub ЗаполнитьОтчетПоКлаймамОН()

    Dim Segodnja As String
    Dim SegodnjaADRES As String
    Dim FSO As Object, fFolders As Object, fFolder As Object
    Dim sFolderName As String
    Dim folderPath As String
    Dim fl As String
    Dim wBook As Workbook
    Dim КоличСканов As Variant
    
' If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
'--------------------------------------------


  
'    On Error GoTo ErrHandler
  
   
   
    Dim pathОтчета As String
    Dim pathОтчета_с_currentmonthname As String
    Dim currentmonthname As String
    
    currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
    pathОтчета = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\Отчет по клаймам за июнь 2025.xlsx" ' полное имя Отчета по клаймам
    pathОтчета_с_currentmonthname = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\" & "Отчет по клаймам за " & currentmonthname & " 2025.xlsx"
'     currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
'    pathОтчета = "C:\Users\Хозяин\Desktop\Отчет по клаймам за июнь 2025.xlsx" ' полное имя Отчета по клаймам
'    pathОтчета_с_currentmonthname = "C:\Users\Хозяин\Desktop\" & "Отчет по клаймам за " & currentmonthname & " 2025.xlsx"
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
        .Height = 645 ' ВЫСОТА окна
        .Left = 0     ' ЛЕВАЯ точка
        .Top = 0    ' ВЕРХНЯЯ точка
    End With
    Worksheets("отчет за день").Select
                        Application.Wait (Now + TimeValue("0:00:02"))
                        Application.DisplayAlerts = False
                        Application.ScreenUpdating = False
    
'   снимаю фильтры на активном листе:
        If wBook.ActiveSheet.FilterMode = True Then
             wBook.ActiveSheet.ShowAllData
        End If
'   Загрузка ДиапазонКлаймов в указанный диапазон:
    Dim iLastRow As Long
    iLastRow = Cells(Rows.Count, 4).End(xlUp).Row
                Cells(iLastRow + 1, 4).Select ' для выделения ячейки, находящейся в последней строке и втором столбце, после последней заполненной ячейки
'                ActiveSheet.Paste ' вставка скопированного диапазона из листа Клаймы2
        ДиапазонКлаймов.Copy  'Вставляем диапазон в выделенную ячейку во второй книге
        ActiveSheet.Paste  'Вставляем значения диапазона в активную ячейку
'        Application.CutCopyMode = False 'Отключаем режим копирования
    Workbooks("Архив. Поиск первички  НСК.xlsx").Activate
    
'    MsgBox КоличествоЗапросов
    Sheets("4692").Range("A1").Select
'If MsgBox("Дальше, удаляю Sheets(Клаймы2)?", vbYesNo) = vbNo Then Exit Sub
    Application.DisplayAlerts = False
    Workbooks("Архив. Поиск первички  НСК.xlsx").Sheets("Клаймы2").Delete '   удаляю Sheets(Клаймы2)
    Workbooks("Отчет по клаймам за сентябрь 2025.xlsx").Activate
    Sheets("отчет за день").Range("A1").Select
    Application.DisplayAlerts = True
'If MsgBox("Вставить ФИО?", vbYesNo) = vbNo Then Exit Sub
    
                iLastRow = Cells(Rows.Count, 2).End(xlUp).Row
                Cells(iLastRow + 1, 2).Select ' для выделения ячейки, находящейся в последней строке и втором столбце, после последней заполненной ячейки
' If MsgBox("Активна ячейка второго столбца?", vbYesNo) = vbNo Then Exit Sub
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
'         Case "Хозяин" ' — ЗЗЗЗЗЗЗЗЗЗЗЗЗЗЗЗАМЕНА !!!
            Cells(iLastRow + 1, 2) = "Лазарев Сергей Александрович"
        Case "g.sidorov"
              Cells(iLastRow + 1, 2) = "Гоша Сидоров"
        End Select
        
'      If MsgBox("Вставилось ФИО? Дальше?", vbYesNo) = vbNo Then Exit Sub
        
        Cells(iLastRow + 1, 3).Select ' — выбор ячейки, расположенной в строке, следующей за последней, и в 3-м столбце.
        
        With Selection.Validation '   в выбранной ячейке создать выпадающий список и выбрать нужное значение для неё
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=тариф!$A$2:$A$10"
        End With
              Cells(iLastRow + 1, 3) = "Подготовка ответа на электронные запросы"
              Cells(iLastRow + 1, 6).Value = Date ' СЕГОДНЯ
     
'     If MsgBox("Протянуть 'Лазарев Сергей Александрович' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
     Cells(ActiveCell.Row, 2).Select 'перейти во второй столбец этой строки
'     If MsgBox("Протянуть 'Лазарев Сергей Александрович' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
     ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличествоЗапросов - 1), ActiveCell.Column)), Type:=xlFillDefault

'     If MsgBox("Протянуть "Подготовка ответа на электронные запросы....' вниз на (КоличСканов - КоличествоБумЗапросов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
     Cells(ActiveCell.Row, 3).Select 'перейти во третий столбец этой строки
     ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличествоЗапросов - 1), ActiveCell.Column)), Type:=xlFillDefault

     Cells(ActiveCell.Row, 6).Select 'перейти в шестой столбец этой строки
'     If MsgBox("Протянуть сегодня вниз на (КоличСканов-1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
     ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличествоЗапросов - 1), ActiveCell.Column)), Type:=xlFillCopy

 Cells(iLastRow + 1, 3).Select
 Cells(iLastRow + 1, 3) = "Подготовка ответа на бумажные запросы"
 ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличествоБумЗапросов - 1), ActiveCell.Column)), Type:=xlFillCopy

     Cells(iLastRow + 1, 4).Select 'перейти в четвертый  столбец этой строки
     Range(Selection, Selection.End(xlDown)).Select   'Выделяю всё вниз от этих ячеек включительно
   
                        Application.ScreenUpdating = True
                        Application.DisplayAlerts = True
    
    Set ДиапазонКлаймов = Nothing
    
'   Подсчитать значения в текущем выделенном диапазоне
    Dim rowCount As Long
    rowCount = Selection.Count
    MsgBox "Количество значений в выбранном диапазоне: " & rowCount
    
'    If MsgBox("Сохранить файл с внесёнными данными?", vbYesNo) <> vbYes Then GoTo Skip
'    'сохранить текущую книгу
'    ActiveWorkbook.Save
'Skip:
    'Закрыть книгу
'    If MsgBox("Закрыть ""Отчет по клаймам""?", vbYesNo) <> vbYes Then Exit Sub
'    ActiveWorkbook.Close
    
'    Exit Sub
'ErrHandler:
'    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Ошибка выполнения"
'   On Error GoTo 0
'   Exit Sub
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








