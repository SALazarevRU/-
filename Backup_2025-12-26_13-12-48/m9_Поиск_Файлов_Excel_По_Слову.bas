Attribute VB_Name = "m9_Поиск_Файлов_Excel_По_Слову"
Option Explicit

Sub ПоискФайловExcelПоСлову() 'Поиск И отображение списка файлов Excel  FindAndListExcelFiles

    Dim fso As Object, objFolder As Object, objFile As Object
    Dim FilePath As String
    Dim i As Long, j As Long ' Добавлена переменная j для итерации по листам
    Dim wb As Workbook, currentWb As Workbook
    Dim ws As Worksheet, targetWs As Worksheet
    Dim HeaderRange As Range, PersonIDHeader As Range
    Dim LastRowA As Long, LastRowB As Long, LastRowC As Long
    Dim FolderPath As String
    Dim FoundOnSheet As String ' Добавлена переменная для хранения имени листа
    
    ' Создаю объекты для работы с файловой системой
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set currentWb = ActiveWorkbook
    
    ' ЯВНО УКАЗЫВАЮ ЛИСТ
    On Error Resume Next ' Обработка ошибки, если лист не существует
    Set targetWs = currentWb.Sheets("Поиск")
    On Error GoTo 0
    
    ' Проверяю, удалось ли получить ссылку на лист
    If targetWs Is Nothing Then
        MsgBox "Лист 'Поиск' не найден! Укажите существующий лист или создайте его."
        Exit Sub
    End If
    
    ' Открываю диалоговое окно выбора каталога
    With Application.FileDialog(msoFileDialogFolderPicker)
                    .title = "Выберите каталог"
                    .AllowMultiSelect = False
        If .Show = -1 Then ' Если пользователь выбрал каталог
            FolderPath = .SelectedItems(1)
        Else ' Если пользователь отменил выбор
            MsgBox "Выбор каталога отменен."
            Exit Sub
        End If
    End With
    
    ' Получаю объект папки
    Set objFolder = fso.GetFolder(FolderPath)
    
    ' Определяю последнюю строку в столбцах A, B и C
    LastRowA = targetWs.Cells(targetWs.Rows.count, "A").End(xlUp).Row
    LastRowB = targetWs.Cells(targetWs.Rows.count, "B").End(xlUp).Row
    LastRowC = targetWs.Cells(targetWs.Rows.count, "C").End(xlUp).Row
    
    ' Увеличиваю счетчики, если столбцы не пустые, иначе начинаю с первой строки
    If LastRowA > 1 Then i = LastRowA + 1 Else i = 1
    If LastRowB > 1 Then j = LastRowB + 1 Else j = 1
    If LastRowC > 1 Then LastRowC = LastRowC + 1
    
    ' Перебираю все файлы в выбранной папке
    For Each objFile In objFolder.Files
    ' Проверяю, является ли файл файлом Excel
    If Right(objFile.Name, 4) = ".xls" Or Right(objFile.Name, 5) = ".xlsx" Or Right(objFile.Name, 5) = ".xlsm" Then
    FilePath = objFile.Path
    
    ' Открываю книгу Excel
    Set wb = Workbooks.Open(FilePath, ReadOnly:=True)
    
    ' Перебираю все листы в книге
    For Each ws In wb.Sheets
    On Error Resume Next
    Set HeaderRange = ws.Rows(1) ' Поиск в первой строке каждого листа
    Set PersonIDHeader = HeaderRange.Find("PersonID", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    
    ' Проверяю, найден ли заголовок "PersonID"
    If Not PersonIDHeader Is Nothing Then
        ' Записываю имя каталога в столбец A
        targetWs.Cells(i, 1).Value = FolderPath
        ' Записываю полное имя файла в столбец B
        targetWs.Cells(i, 2).Value = objFile.Path
        ' Записываю имя листа в столбец C
        targetWs.Cells(i, 3).Value = ws.Name
        i = i + 1
        End If
        Next ws
        
        ' Закрываю книгу, не сохраняя изменения
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    DoEvents ' Даю Excel возможность обработать события
    Next objFile
    
    ' Проверяю, были ли найдены файлы с заголовком "PersonID"
    If i = 1 Then
        MsgBox "В данном каталоге книг с заголовком столбца PersonID не найдено"
    End If
    
    'Очищаю объекты
    Set fso = Nothing
    Set objFolder = Nothing
    Set objFile = Nothing
    Set HeaderRange = Nothing
    Set PersonIDHeader = Nothing
    Set ws = Nothing
    Set targetWs = Nothing

End Sub

Sub ОткрытьКнигуИЛист()
    Dim ПутьКФайлу As String
    Dim ИмяЛиста As String
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ПутьКФайлу = ActiveCell.Value ' Получаю полный путь к файлу из активной ячейки
    
    ИмяЛиста = ActiveCell.Offset(0, 1).Value ' Получаю имя листа из ячейки справа от активной
    
    On Error GoTo ОбработкаОшибок ' Обработка ошибок
    
    Set wb = Workbooks.Open(ПутьКФайлу) ' Открываю книгу
    
    ' Активирую указанный лист
    Set ws = wb.Sheets(ИмяЛиста)
    ws.Activate
    
    
    Exit Sub ' Выход из процедуры, если все прошло успешно
    
ОбработкаОшибок:
    ' Обрабатываю возможные ошибки
    If Err.Number = 9 Then
    MsgBox "Лист с именем '" & ИмяЛиста & "' не найден в книге.", vbCritical, "Ошибка"
    ElseIf Err.Number = 1004 Then
    MsgBox "Не удалось открыть книгу по указанному пути: " & ПутьКФайлу & ". Убедитесь, что путь указан верно и файл существует.", vbCritical, "Ошибка"
    Else
    MsgBox "Произошла ошибка: " & Err.Description, vbCritical, "Ошибка"
    End If
    
    ' Очищаю объектные переменные
    Set wb = Nothing
    Set ws = Nothing

End Sub

Sub ОткрытьКнигу333()
    Dim ПутьКФайлу As String
    Dim wb As Workbook
    ПутьКФайлу = "C:\Users\Хозяин\Desktop\PersonID - копия\Книга6.xlsx" ' Полный путь к файлу
   On Error GoTo ОбработкаОшибок ' Обработка ошибок
    Set wb = Workbooks.Open(ПутьКФайлу)  ' Открываю книгу
    Exit Sub  ' Выход из процедуры, если все прошло успешно
ОбработкаОшибок: ' Обрабатываю возможные ошибки
        If Err.Number = 9 Then
            MsgBox "Не удалось открыть книгу по указанному пути: " & ПутьКФайлу & ". Убедитесь, что путь указан верно и файл существует.", vbCritical, "Ошибка"
        Else
            MsgBox "Произошла ошибка: " & Err.Description, vbCritical, "Ошибка"
        End If
    Set wb = Nothing ' Очищаю объектные переменные
End Sub

Sub ОткрытьКнигу444()
    Dim ПутьКФайлу As String
    Dim wb As Workbook
    ПутьКФайлу = "C:\Users\Хозяин\Desktop\PersonID - копия\Книга6.xlsx" ' Полный путь к файлу
    Set wb = Workbooks.Open(ПутьКФайлу)  ' Открываю книгу
    Set wb = Nothing ' Очищаю объектные переменные
End Sub


Sub FindAndOpenFile()

    Dim FilePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Integer
    Dim LastCol As Long
    Dim HeaderValue As String
    Dim Answer As VbMsgBoxResult
    
    ' 1. Выбор файла
    FilePath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", title:="Выберите Excel-файл")
    
    If FilePath = "False" Then ' Если пользователь отменил выбор
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(FilePath) ' 2. Открытие файла
    
    ' 3. Перебор листов
    For Each ws In wb.Worksheets
    
    ' 4. Поиск столбца на текущем листе
    LastCol = ws.Cells(1, Columns.count).End(xlToLeft).Column ' Определение последней колонки с данными
    
    For i = 1 To LastCol
    HeaderValue = ws.Cells(1, i).Value ' Получение значения заголовка столбца
    
        If InStr(1, HeaderValue, "PersonID", vbTextCompare) > 0 Then ' Проверка на наличие "PersonID" (нечувствительно к регистру)
        
            ' 5. Уведомление и выбор
            Answer = MsgBox("Слово найдено, открыть файл?", vbYesNo, "Найдено PersonID")
            
            If Answer = vbYes Then
                wb.Activate ' Активирую открытый файл (если он уже открыт)
                Exit Sub ' Завершаю работу макроса
            Else
            
                Exit Sub ' Завершаю работу макроса
            End If
        
        End If
    
    Next i
    
    Next ws
    
    MsgBox "Столбец PersonID не найден ни на одном листе.", vbInformation
    wb.Close
End Sub

