Attribute VB_Name = "m9_Поиск_файлов_эксель_с_словом"
Sub ПоискФайловExcelСоСловом()
    ' Объявление переменных
    Dim fso As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim wbk As Workbook
    Dim wks As Worksheet
    Dim rFound As Range
    Dim strFirstAddress As String
    Dim strSearch As String
    Dim folderPath As String
    Dim wOut As Worksheet
    Dim lRow As Long
    
    ' Запрос строки для поиска
    strSearch = InputBox("Введите строку для поиска:", "Поиск текста")
    If strSearch = "" Then
        MsgBox "Строка поиска не введена.", vbExclamation
        Exit Sub
    End If
    
    ' Диалоговое окно выбора каталога
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Выберите каталог"
        .AllowMultiSelect = False
        If .Show = -1 Then ' Если пользователь выбрал каталог
            folderPath = .SelectedItems(1)
        Else ' Если пользователь отменил выбор
            MsgBox "Выбор каталога отменен."
            Exit Sub
        End If
    End With
    
    ' Создание объекта FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFolder = fso.GetFolder(folderPath)
    
    ' Создание нового листа для вывода результатов
    Set wOut = Worksheets.Add
    ' Присваиваем листу имя
    wOut.Name = "Поиск по слову"
    lRow = 1
    With wOut
        .Cells(lRow, 1) = "Файл"
        .Cells(lRow, 2) = "Лист"
        .Cells(lRow, 3) = "Ячейка"
        .Cells(lRow, 4) = "Текст"
    End With
    
    ' Перебор всех файлов в папке
    For Each objFile In objFolder.Files
        ' Проверка расширения файла (xls, xlsx, xlsm)
        If LCase(Right(objFile.Name, 4)) = ".xls" Or _
           LCase(Right(objFile.Name, 5)) = ".xlsx" Or _
           LCase(Right(objFile.Name, 5)) = ".xlsm" Then
            
            On Error Resume Next
            Set wbk = Workbooks.Open(FileName:=objFile.Path, _
                                      UpdateLinks:=0, _
                                      ReadOnly:=True, _
                                      AddToMRU:=False)
            On Error GoTo 0
            
            If Not wbk Is Nothing Then
                ' Перебор всех листов в файле
                For Each wks In wbk.Worksheets
                    Set rFound = wks.UsedRange.Find(What:=strSearch, _
                                                   LookIn:=xlValues, _
                                                   LookAt:=xlPart, _
                                                   SearchOrder:=xlByRows)
                    
                    If Not rFound Is Nothing Then
                        strFirstAddress = rFound.Address
                        Do
                            lRow = lRow + 1
                            wOut.Cells(lRow, 1) = objFile.Name
                            wOut.Cells(lRow, 2) = wks.Name
                            wOut.Cells(lRow, 3) = rFound.Address
                            wOut.Cells(lRow, 4) = rFound.Value
                            
                            Set rFound = wks.Cells.FindNext(After:=rFound)
                            If rFound Is Nothing Then Exit Do
                        Loop While rFound.Address <> strFirstAddress
                    End If
                Next wks
                
                wbk.Close False ' Закрытие файла без сохранения
            End If
        End If
    Next objFile

    ' Форматирование результата
    With wOut
        .Columns("A:D").EntireColumn.AutoFit
        .Rows(1).Font.Bold = True
    End With

    MsgBox "Поиск завершен. Найдено " & lRow - 1 & " совпадений.", vbInformation
End Sub


