Attribute VB_Name = "m9_Заполнить_Отчет_ФКБ"
Option Explicit
Option Compare Text

Sub ЗаполнитьОтчетПоСканированию(control As IRibbonControl) '   control As IRibbonControl
    Dim КоличСканов As Variant
    Dim Segodnja As String
    Dim SegodnjaADRES As String
    Dim fso As Object, fFolders As Object, fFolder As Object
    Dim sFolderName As String
    Dim FolderPath As String
    Dim fl As String
    Dim wBook As Workbook
    
    On Error GoTo ErrHandler
'--------------------------------------------
Application.DisplayAlerts = False

    '    FolderPath = "C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\СКАНЫ_за_ВЧЕРА\"
    FolderPath = PapkaSkanov
    ' получаем характеристики папки
    Set fso = CreateObject("Scripting.FileSystemObject")
    КоличСканов = fso.GetFolder(FolderPath).Subfolders.Count
'    If MsgBox("Как бэ сегодня папок со сканами = " & КоличСканов & " Открыть файл Отчёта?", vbYesNo) <> vbYes Then Exit Sub
    

    MsgBox "Как бэ сегодня папок со сканами = " & КоличСканов
    
    
    fl = IsBookOpen("C:\Users\Хозяин\Desktop\Отчет по клаймам за декабрь 2025.xlsx")
        If fl Then ' если уже был открыт мною
            Set wBook = Workbooks("Отчет по клаймам за май 2025.xlsx")
            wBook.Windows(1).Activate 'вытягиваем на первый план
        Else
            Set wBook = Workbooks.Open(FileName:="C:\Users\Хозяин\Desktop\Отчет по клаймам за декабрь 2025.xlsx")  ' открыл
        End If
    
    With Application  ' размещаю окно  в координатах:
        .WindowState = xlNormal
        .Width = 1420 ' ШИРИНА окна
        .Height = 445 ' ВЫСОТА окна
        .Left = 0     ' ЛЕВАЯ точка
        .Top = 0    ' ВЕРХНЯЯ точка
    End With
     Worksheets("отчет за день").Select
                         Application.ScreenUpdating = False  ' можно отрубить отображение процессов на экране
'   снимаю фильтры на активном листе:
    If wBook.ActiveSheet.FilterMode = True Then
         wBook.ActiveSheet.ShowAllData
    End If
'   Загрузка списка папок в указанный диапазон:
     Dim i           As Long
'    On Error GoTo ErrHandler
 
    Dim L           As String
    L = "C:\Users\Хозяин\Desktop\Сканы АБВ"
 
    Dim coll        As Collection
    Set coll = SubFoldersCollection(L)
 
    With ActiveWorkbook.Worksheets("отчет за день")    ' замените на конкретное имя вашего листа
 
        Dim nextRow As Long
        nextRow = .Cells(.Rows.Count, "C").End(xlUp).Row + 1
        If .Cells(1, "C").Value = "" Then nextRow = 1
 
        For i = 1 To coll.Count
            .Cells(nextRow, 3).Value = coll(i)
            nextRow = nextRow + 1
        Next i
 
    End With
'    Exit Sub
 
'ErrHandler:
'    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Ошибка выполнения"
'    On Error GoTo 0
    
'  Sleep 400
  '     If MsgBox("Копировать 'Лазарев Сергей Александрович'?", vbYesNo) <> vbYes Then Exit Sub
    Worksheets("отчет за день").Range("A142").Copy
'         If MsgBox(Range("A142").Value, vbYesNo) <> vbYes Then Exit Sub
        Dim iLastRow As Long
        iLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Cells(iLastRow + 1, 1).Select
'     If MsgBox("Вставить 'Лазарев Сергей Александрович'?", vbYesNo) <> vbYes Then Exit Sub
ActiveSheet.Paste
'     If MsgBox("Протянуть 'Лазарев Сергей Александрович' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillDefault
'  If MsgBox("Копировать 'Изъятие документов из....'?", vbYesNo) <> vbYes Then Exit Sub
Worksheets("отчет за день").Range("B142").Copy
'     If MsgBox("Выбрать первую пустую ячейку столбца B?", vbYesNo) <> vbYes Then Exit Sub
'Range("A" & Cells(Rows.Count, 3).End(xlUp).Row + 1).Select
        iLastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Cells(iLastRow + 1, 2).Select
'     If MsgBox("Вставить 'Изъятие документов из....'?", vbYesNo) <> vbYes Then Exit Sub
ActiveSheet.Paste
'     If MsgBox("Протянуть 'Изъятие документов из....' вниз на (КоличСканов - 1) ячеек?", vbYesNo) <> vbYes Then Exit Sub
ActiveCell.AutoFill Destination:=Range(ActiveCell, Cells(ActiveCell.Row + (КоличСканов - 1), ActiveCell.Column)), Type:=xlFillDefault
'    Range("D2:D10222").NumberFormat = "mm/dd/yyyy"

' If MsgBox("Выбрать первую пустую ячейку столбца 4?", vbYesNo) <> vbYes Then Exit Sub
    Range("D" & Cells(Rows.Count, 4).End(xlUp).Row + 1).Select
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
 


Sub Макрос1()
Dim iLastRow As Long
iLastRow = Cells(Rows.Count, 2).End(xlUp).Row
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




