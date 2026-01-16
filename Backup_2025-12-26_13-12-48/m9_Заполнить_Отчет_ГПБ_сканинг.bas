Attribute VB_Name = "m9_Заполнить_Отчет_ГПБ_сканинг"
Option Explicit
Option Compare Text
Public PapkaSkanov, lf As String ' перменная становится доступной для использования в любом модуле и любой процедуре проекта

Sub ЗаполнитьОтчетПоГазпрому(control As IRibbonControl)
'Sub ЗаполнитьОтчетПоГазпрому() '
    '    Dim КоличСканов As Variant
    Dim Segodnja As String
    Dim SegodnjaADRES As String
    Dim fso As Object, fFolders As Object, fFolder As Object
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
    Set fso = CreateObject("Scripting.FileSystemObject")
    КоличСканов = fso.GetFolder(FolderPath).Subfolders.count
         MsgBox "Выбран каталог " & PapkaSkanov & vbNewLine & vbNewLine & "В этом каталоге " & КоличСканов & " папок со сканами" ' = " & КоличСканов
         
    Dim pathОтчета As String
    Dim pathОтчета_с_currentmonthname As String
    Dim currentmonthname As String
    
    currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
    pathОтчета = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\Отчет по клаймам за июнь 2025.xlsx" ' полное имя Отчета по клаймам
    pathОтчета_с_currentmonthname = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\" & "Отчет по клаймам за " & currentmonthname & " 2025.xlsx"
    
    f1_Ожидайте.Show
    
    Application.ScreenUpdating = False  'ОКЛЮЧАЮ ОБНОВЛЕНИЕ ЭКРАНА
    
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
        .Width = 1442 ' ШИРИНА окна
        .Height = 790 ' ВЫСОТА окна
        .Left = 0     ' ЛЕВАЯ точка
        .Top = 0    ' ВЕРХНЯЯ точка
    End With
    Worksheets("отчет за день").Select
    
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
        nextRow = .Cells(.Rows.count, "D").End(xlUp).Row + 1
        If .Cells(1, "D").Value = "" Then nextRow = 1
 
        For i = 1 To coll.count
            .Cells(nextRow, 4).Value = coll(i)
            nextRow = nextRow + 1
        Next i
 
    End With
    
    Application.Wait (Now + TimeValue("0:00:01"))  ' Пауза
    'If MsgBox("Дальше?", vbYesNo) = vbNo Then Exit Sub
    Dim iLastRow As Long
                iLastRow = Cells(Rows.count, 2).End(xlUp).Row
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
              Cells(iLastRow + 1, 2) = "Лященко Альфия"
        End Select
     
'     If MsgBox("Протянуть 'Лазарев Сергей Александрович' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub

     ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillDefault


    Dim r1 As Range, r2 As Range
    Worksheets("отчет за день").Select '  СУКА ТУТ ЗАСАДА Т.К. ЕСЛИ ТЫ ПЕРВЫЙ ЗАПОЛНЯЕШЬ НОВЫЙ ФАЙЛ 1-ГО ЧИСЛА ТО НЕ НАЙДЕШЬ "ИЗЪЯТИЕ..." В СТОЛБЦЕ !!!
    Set r1 = Columns("C:C").Find("Изъятие документов из поступивших от контрагента досье, согласно матрицы компании, их сканирование и переименование СРЕДНЕЕ ДОСЬЕ")
    If Not r1 Is Nothing Then r1.Copy

''  If MsgBox("Копировать 'Изъятие документов из....'?", vbYesNo) <> vbYes Then Exit Sub
'Worksheets("отчет за день").Range("B760").Copy


'     If MsgBox("Выбрать первую пустую ячейку столбца B?", vbYesNo) <> vbYes Then Exit Sub
'Range("A" & Cells(Rows.Count, 3).End(xlUp).Row + 1).Select
        iLastRow = Cells(Rows.count, 3).End(xlUp).Row
        Cells(iLastRow + 1, 3).Select
'     If MsgBox("Вставить 'Изъятие документов из....'?", vbYesNo) <> vbYes Then Exit Sub
ActiveSheet.Paste
'     If MsgBox("Протянуть 'Изъятие документов из....' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillDefault
'    Range("D2:D10222").NumberFormat = "mm/dd/yyyy"

' If MsgBox("Выбрать первую пустую ячейку столбца 4?", vbYesNo) <> vbYes Then Exit Sub
    Range("F" & Cells(Rows.count, 6).End(xlUp).Row + 1).Select
'     If MsgBox("Вставить сегодняшнюю дату?", vbYesNo) <> vbYes Then Exit Sub
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Segodnja = ActiveCell.Value
    
'    If MsgBox("Segodnja= " & Segodnja, vbYesNo) <> vbYes Then Exit Sub
    SegodnjaADRES = ActiveCell.Address
'    If MsgBox("SegodnjaADRES= " & SegodnjaADRES, vbYesNo) <> vbYes Then Exit Sub
'     If MsgBox("Заменить её на значение?", vbYesNo) <> vbYes Then Exit Sub
    ActiveCell = ActiveCell.Value
    
'    If MsgBox("Протянуть вниз на (КоличСканов-1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
    ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillCopy
'    ActiveCell.Offset(0, 1).Activate 'Перехожу на одну ячейку правее активной и активируем ее.



'    II - Вторая стадия:-------------------------------------------------------------------------------------

'   Загрузка списка папок в указанный диапазон:
     
    Set coll = SubFoldersCollection(L)
 
    With ActiveWorkbook.Worksheets("отчет за день")    ' замените на конкретное имя вашего листа
 
       
        nextRow = .Cells(.Rows.count, "D").End(xlUp).Row + 1
        If .Cells(1, "D").Value = "" Then nextRow = 1
 
        For i = 1 To coll.count
            .Cells(nextRow, 4).Value = coll(i)
            nextRow = nextRow + 1
        Next i
 
    End With
     Worksheets("отчет за день").Range("B2132").Copy
'         If MsgBox(Range("A142").Value, vbYesNo) <> vbYes Then Exit Sub
'        Dim iLastRow As Long
        iLastRow = Cells(Rows.count, 2).End(xlUp).Row
        Cells(iLastRow + 1, 2).Select
'     If MsgBox("Вставить 'Лазарев Сергей Александрович'?", vbYesNo) <> vbYes Then Exit Sub
Cells(iLastRow + 1, 2).Value = "Лазарев Сергей Александрович"
'     If MsgBox("Протянуть 'Лазарев Сергей Александрович' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillDefault


Set r1 = Columns("C:C").Find("Сверка документов от контрагентов по кол-ву должников")
    If Not r1 Is Nothing Then r1.Copy
    
''  If MsgBox("Копировать 'Сверка документов от ....'?", vbYesNo) <> vbYes Then Exit Sub
'Worksheets("отчет за день").Range("B1148").Copy

'     If MsgBox("Выбрать первую пустую ячейку столбца B?", vbYesNo) <> vbYes Then Exit Sub
'Range("A" & Cells(Rows.Count, 3).End(xlUp).Row + 1).Select
        iLastRow = Cells(Rows.count, 3).End(xlUp).Row
        Cells(iLastRow + 1, 3).Select
'     If MsgBox("Вставить 'Сверка документов от....'?", vbYesNo) <> vbYes Then Exit Sub
ActiveSheet.Paste
'     If MsgBox("Протянуть 'Сверка документов от....' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillDefault
'    Range("D2:D10222").NumberFormat = "mm/dd/yyyy"

' If MsgBox("Выбрать первую пустую ячейку столбца 4?", vbYesNo) <> vbYes Then Exit Sub
    Range("F" & Cells(Rows.count, 6).End(xlUp).Row + 1).Select
'     If MsgBox("Вставить сегодняшнюю дату?", vbYesNo) <> vbYes Then Exit Sub
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Segodnja = ActiveCell.Value
    
'    If MsgBox("Segodnja= " & Segodnja, vbYesNo) <> vbYes Then Exit Sub
    SegodnjaADRES = ActiveCell.Address
'    If MsgBox("SegodnjaADRES= " & SegodnjaADRES, vbYesNo) <> vbYes Then Exit Sub
'     If MsgBox("Заменить её на значение?", vbYesNo) <> vbYes Then Exit Sub
    ActiveCell = ActiveCell.Value
    
'    If MsgBox("Протянуть вниз на (КоличСканов-1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
    ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillCopy
'    ActiveCell.Offset(0, 1).Activate 'Перехожу на одну ячейку правее активной и активируем ее.
    ActiveCell.Select
    Range(Selection, Selection.End(xlDown)).Select   'Выделяю всё вниз от этих ячеек включительно
   
    Application.ScreenUpdating = True
'   Подсчитать значения в текущем выделенном диапазоне
    Dim rowCount As Long
    rowCount = Selection.count
    
    Unload f1_Ожидайте
    
    Application.ScreenUpdating = True   'ВКЛЮЧАЮ ОБНОВЛЕНИЕ ЭКРАНА
    
    MsgBox "Количество значений в выбранном диапазоне: " & rowCount
    
    If MsgBox("Заполнить данными иное время?", vbYesNo) <> vbYes Then GoTo Skip_0
    
    Worksheets("иное время").Select
        iLastRow = Cells(Rows.count, 2).End(xlUp).Row
           Cells(iLastRow + 1, 2).Select
    With Selection.Validation '   в выбранной ячейке создать выпадающий список и выбрать нужное значение для неё
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Ващенко Оксана Васильевна,Лазарев Сергей Александрович"
            End With
            
    Application.Wait (Now + TimeValue("0:00:01")) ' Пауза
    
'    Dim username As String ' Тут второй раз уже объявил тип переменной, не нужно.
    username = Environ("UserName")  ' Получаем имя пользователя компьютера.
        
    Select Case Environ("UserName")
        Case "v.petrov"
              Cells(iLastRow + 1, 2) = "Ващенко Оксана Васильевна"
        Case "s.lazarev"
            Cells(iLastRow + 1, 2) = "Лазарев Сергей Александрович"
        Case "g.sidorov"
              Cells(iLastRow + 1, 2) = "Гоша Сидоров"
        End Select
        
    Cells(iLastRow + 1, 3).Value = "Иная работа" ' — выбор ячейки, расположенной в строке, следующей за последней, и в 3-м столбце.
    Cells(iLastRow + 1, 6).Value = Date ' СЕГОДНЯ
    Cells(iLastRow + 1, 7).Value = "задача 45712, раскладка КД"
    Cells(iLastRow + 1, 4).Select
    
    f1_Цифровая_клавиатура.Show 1
    
Skip_0:
    
'    If MsgBox("Сохранить файл с внесёнными данными?", vbYesNo) <> vbYes Then GoTo Skip
    'сохранить текущую книгу
'    ActiveWorkbook.Save
Skip:
    'Закрыть книгу
'    If MsgBox("Закрыть ""Отчет по клаймам""?", vbYesNo) <> vbYes Then Exit Sub
'    ActiveWorkbook.Close
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

Sub Макрос1()
Dim iLastRow As Long
iLastRow = Cells(Rows.count, 2).End(xlUp).Row
Cells(iLastRow + 1, 4).Select
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
    Dim folder      As Object
    On Error GoTo ErrHandler
 
    Set SubFoldersCollection = New Collection
 
    Dim fso         As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
 
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    Set curfold = fso.GetFolder(FolderPath)
 
    If Not curfold Is Nothing Then
 
        For Each folder In curfold.Subfolders
 
            If folder.Path Like FolderPath & Mask Then
 
                ' только имена подпапок
                SubFoldersCollection.Add folder.Name
 
                '                ' полные пути к подпапкам с именами подпапок
                '                SubFoldersCollection.Add folder.Path & "\"
            End If
 
        Next folder
 
    End If
 
    Set fso = Nothing
    Set curfold = Nothing
    Exit Function
 
ErrHandler:
    MsgBox "Ошибка при получении подпапок: " & Err.Description, vbCritical, "Ошибка"
    Set SubFoldersCollection = Nothing
    Set fso = Nothing
    Set curfold = Nothing
End Function

