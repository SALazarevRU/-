Attribute VB_Name = "m9_Команды_меню"
Sub СкроллВверх(control As IRibbonControl) ' Шобы снизу таблы вверх подняццо
'ActiveWindow.ScrollColumn = 299
'ActiveWindow.ScrollRow = 3
ActiveWindow.ScrollRow = 1
Application.Wait Now + TimeSerial(0, 0, 0.5)
End Sub

' Public позволяет сделать подпрограмму видимой во всех модулях рабочей книги.
Sub checkbox01_startup(control As IRibbonControl, ByRef returnedVal) 'ByRef-может изменить переменную
    If ActiveWindow.DisplayGridlines = True Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub


Sub checkbox01_clicked(control As IRibbonControl, pressed As Boolean)
    Select Case pressed
        Case True
            ActiveWindow.DisplayGridlines = True
        Case False
            ActiveWindow.DisplayGridlines = False
    End Select
End Sub
        
'gRibbon.ActivateTabQ "Tab1", "somename" ' команда используется для автоматического открытия пользовательской вкладки («Валидация ПД») при запуске документа.
'В аргументах команды указывается имя вкладки («Tab1») и название пространства имён («somename»). Оба аргумента обязательны.


'MsgBox TypeName(sName)

Sub ClearClipboard() ' очистить буфер экселя
    Dim DataObj As New MSForms.DataObject
    DataObj.SetText ""
    DataObj.PutInClipboard
End Sub

Sub ЗаписатьВClipboard() ' записать в буфер экселя
   With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
   .SetText Err.Description: .PutInClipboard
End Sub



                                           '1_часть Private Sub Сменить_Обои_рабочего_стола() это функ Public Declare PtrSafe Function SystemParametersInfo и конст Public Const SPI_SETDESKWALLPAPER = 20 нах-ся в модуле m_Все_Declare_PtrSafe
Sub SetWallpaper(file As String)   '2_часть Private Sub Сменить_Обои_рабочего_стола()
SystemParametersInfo SPI_SETDESKWALLPAPER, 0, ByVal file, True
End Sub
 
Sub Смена_Обоев(control As IRibbonControl)  '3_часть Private Sub Сменить_Обои_рабочего_стола()
'SetWallpaper "C:\Users\Хозяин\Desktop\123.bmp"
SetWallpaper "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\CustomUI_LS\Черный_Фон3.jpg"
End Sub

Private Sub УдалитьСкрытыеЛисты()
Dim Sh As Object
    For Each Sh In ActiveWorkbook.Sheets
    Application.DisplayAlerts = False  'Отключаем запрос подтверждения на удаление
        If Sh.Visible <> xlSheetVisible Then Sh.Delete 'Если лист скрытый то удаляем его
        Application.DisplayAlerts = True
    Next
MsgBox "Скрытые листы удалены.", vbInformation, "Отчёт"
End Sub

Public Sub Фильтры_все_очистить()  ' ОЧИСТИТЬ а не УДАЛИТЬ!
    If ActiveSheet.FilterMode = True Then
      ActiveSheet.ShowAllData
    End If
    ActiveSheet.Cells.Rows.Hidden = False
   
End Sub


Sub ПоказатьСписокПапокВДиректории_1() ' Список папок (подпапок) в Директории
    Dim pDialog As FileDialog, pFolder As Object
    Dim fso As Object, nextFolder As Object
    Dim folderNames() As String, i As Long
    
    If MsgBox("Список папок в указанной Вами директории будет выгружен в столбец F активного листа, продолжить?", vbYesNo) = vbNo Then Exit Sub
    
    Set pDialog = Application.FileDialog(msoFileDialogFolderPicker)
    pDialog.AllowMultiSelect = False
    
    If pDialog.Show Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set pFolder = fso.GetFolder(pDialog.SelectedItems(1))
        ReDim folderNames(1 To pFolder.Subfolders.count)
        i = 0
        
        If Range("F1") = "" Then
                For Each nextFolder In pFolder.Subfolders
                i = i + 1
                folderNames(i) = nextFolder.Name
                Range("F" & i) = folderNames(i) ' заябись, в разные ячейки...
                Next
            Else
                For Each nextFolder In pFolder.Subfolders
                i = i + 1
                folderNames(i) = nextFolder.Name
                Range("G" & i) = folderNames(i)
                Next
      
'          If Not Range("G1") = "" Then
'                For Each nextFolder In pFolder.Subfolders
'                i = i + 1
'                folderNames(i) = nextFolder.Name
'                Range("H" & i) = folderNames(i)
'                Next
        End If
     End If
   
'       Debug.Print Join(folderNames, vbLf)
'      ActiveSheet.Range("F1").Value = Join(folderNames, vbLf) ' все в одну ячейку...

      
 '      УСЛОВНОЕ ФОРМАТИРОВАНИЕ ТРЕХ СТОЛБЦОВ
    Columns("F:H").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub


Sub ПоказатьСписокПапокВДиректории(control As IRibbonControl)
    Dim pDialog As FileDialog, pFolder As Object
    Dim fso As Object, nextFolder As Object
    Dim folderNames() As String, i As Long
    Dim targetColumn As String
    
    If MsgBox("Список папок в указанной Вами директории будет выгружен в столбец F активного листа, продолжить?", vbYesNo) = vbNo Then Exit Sub
    
    Set pDialog = Application.FileDialog(msoFileDialogFolderPicker)
    pDialog.AllowMultiSelect = False
    
    If pDialog.Show Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pFolder = fso.GetFolder(pDialog.SelectedItems(1))
    ReDim folderNames(1 To pFolder.Subfolders.count)
    i = 0
    
        ' Определяю, в какой столбец вставлять данные
        If WorksheetFunction.CountA(Columns("F")) = 0 Then
           targetColumn = "F"
        ElseIf WorksheetFunction.CountA(Columns("G")) = 0 Then
               targetColumn = "G"
        ElseIf WorksheetFunction.CountA(Columns("H")) = 0 Then
               targetColumn = "H"
        Else
               MsgBox "Столбцы F, G и H заполнены. Некуда вставлять список папок.", vbExclamation
               Exit Sub
        End If
    
    ' Заполняю столбец именами папок
    For Each nextFolder In pFolder.Subfolders
    i = i + 1
    folderNames(i) = nextFolder.Name
    Range(targetColumn & i).Value = folderNames(i)
    Next nextFolder
End If

    'Очищаю переменные из памяти
    Set pDialog = Nothing
    Set fso = Nothing
    Set pFolder = Nothing
    Set nextFolder = Nothing

'   УСЛОВНОЕ ФОРМАТИРОВАНИЕ ТРЕХ СТОЛБЦОВ
    Columns("F:H").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

End Sub



Sub СНЯТЬУСЛОВНОЕФОРМАТИРОВАНИЕ(control As IRibbonControl)
    Cells.FormatConditions.Delete
End Sub

Sub СоздатьТекстовыйФайлВDownloads(control As IRibbonControl) ' ПОИСК ОТЛИЧКА
'Sub СоздатьТекстовыйФайлВDownloads()
Dim fso As Object, ts As Object
Dim strFilePath As String

' Определяем путь к папке Downloads
strFilePath = Environ("USERPROFILE") & "\Downloads\Документы не найдены.txt"

' Создаем объект FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Проверяем, существует ли файл. Если да, удаляем его.
If fso.fileExists(strFilePath) Then
fso.DeleteFile strFilePath
End If

' Создаем текстовый файл
Set ts = fso.CreateTextFile(strFilePath, True) ' True означает перезапись, если файл существует

' (Опционально) Записываем текст в файл
' ts.WriteLine "Это текст, который будет записан в файл."

' Закрываем текстовый файл
ts.Close

' Освобождаем объекты
Set ts = Nothing
Set fso = Nothing

'MsgBox "Текстовый файл 'Новый текстовый файл.txt' создан в папке Downloads!", vbInformation
CreateObject("WScript.Shell").Popup "'Документы не найдены.txt' создан в папке Downloads!" & _
                    vbNewLine & _
                    vbNewLine & " Это окно закроется через 2 секунд", 2, "Сообщение о завершении операции"
 
'Windows("Архив. Поиск первички  НСК.xlsx").Activate ' Активирую книгу. нет.
AppActivate ("Архив. Поиск первички  НСК.xlsx - Excel")  ' Активирую книгу. АКТИВИРУЕТСЯ. (вар-3)
ActiveCell.Select
Cells(ActiveCell.Row, "M").Select  ' переход из ячейки строки в ячейку этой же строки но уже конкретного столбца М
f1_Нет_транша_нет_рко_нет.Show
End Sub




Sub СоздатьТекстовыеФайлыВПустыхПапках() ' ПОИСК ОТЛИЧКА

  Dim fso As Object, folder As Object, subfolder As Object
  Dim FolderPath As String, FilePath As String
  Dim emptyFolderFound As Boolean
  
'  folderPath = "C:\Users\Хозяин\Desktop\Сканы АБВ" ' Указал путь к директории "Сканы"
  FolderPath = "C:\Users\s.lazarev\Downloads" ' Указал путь к директории "Сканы"

  ' Создал объект FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Проверяю, существует ли указанная директория
  If Not fso.FolderExists(FolderPath) Then
    MsgBox "Директория '" & FolderPath & "' не найдена.", vbCritical
    Exit Sub
  End If

  ' Получил объект Folder, представляющий директорию "Downloads"
  Set folder = fso.GetFolder(FolderPath)

  emptyFolderFound = False ' Флаг для отслеживания наличия пустых папок

  ' Перебираю все подпапки в директории "Сканы"
  For Each subfolder In folder.Subfolders
    ' Проверяю, пустая ли папка
    If subfolder.Files.count = 0 And subfolder.Subfolders.count = 0 Then
      ' Создаю путь к новому текстовому файлу
      FilePath = subfolder.Path & "\Новый текстовый документ.txt"

      fso.CreateTextFile FilePath ' Создаю пустой текстовый файл
      emptyFolderFound = True ' Отмечаю, что пустая папка была найдена
    End If
  Next subfolder

  ' Если пустых папок не найдено, выхожу из процедуры
  If Not emptyFolderFound Then
    MsgBox "Пустых папок в директории '" & FolderPath & "' не найдено.", vbInformation
    Exit Sub
  End If

'  MsgBox "Создание текстовых файлов завершено!", vbInformation

  ' Очищаю объекты
  Set fso = Nothing
  Set folder = Nothing
  Set subfolder = Nothing

End Sub

Sub ПереместитьПопдпапкиВ_Количестве(control As IRibbonControl) 'MoveSubfolders
       If MsgBox("Запустить ""Перемещение попдпапок в заданном количестве""?", vbYesNo, "Имя приложения: " & Application.Name) = vbNo Then Exit Sub
  Dim SourcePath As String
  Dim DestinationPath As String
  Dim NumFoldersToMove As Variant
  Dim fso As Object
  Dim SourceFolder As Object
  Dim subfolder As Object
  Dim Subfolders As Object
  Dim i As Integer
  Dim FolderArray() As String
  Dim folderCount As Integer
  Dim j As Integer ' Добавлено объявление переменной j
  Dim Temp As String ' Добавлено объявление переменной Temp

  ' Запрашиваю у пользователя пути к каталогам
  SourcePath = InputBox("Введите путь к каталогу-источнику (A1):", "Путь к Источнику", "C:\Users\s.lazarev\Desktop\ПОИСК первички\СКАНЫ")
  If SourcePath = "" Then Exit Sub ' Если пользователь отменил ввод

  DestinationPath = InputBox("Введите путь к каталогу-назначению (A2):", "Путь к Назначению", "C:\Users\s.lazarev\Desktop\ПОИСК первички\СКАНЫ для подготовки к сливу")
  If DestinationPath = "" Then Exit Sub ' Если пользователь отменил ввод

  ' Запрашиваю у пользователя количество подпапок для перемещения
  Dim Клаймов As Integer
  Клаймов = Worksheets("ппон").Range("D11").Value
  NumFoldersToMove = InputBox("(Клаймов " & Клаймов & "). Введите количество подпапок для перемещения:", "Количество Подпапок")
  If NumFoldersToMove = "" Then Exit Sub     ' Если пользователь отменил ввод
  If Not IsNumeric(NumFoldersToMove) Then
     MsgBox "Введено некорректное значение. Пожалуйста, введите число.", vbCritical
    Exit Sub
  End If

  ' Создаю объект FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Получаю объект папки-источника
  Set SourceFolder = fso.GetFolder(SourcePath)

  ' Получаю коллекцию подпапок
  Set Subfolders = SourceFolder.Subfolders

  ' Определяю размер массива для хранения путей к подпапкам
  ReDim FolderArray(1 To Subfolders.count)
  folderCount = 0

  ' Заполняю массив путями к подпапкам
  For Each subfolder In Subfolders
    folderCount = folderCount + 1
    FolderArray(folderCount) = subfolder.Path
  Next subfolder

  ' Сортирую массив подпапок по дате создания (от новых к старым)
  ' (Для простоты используется простой алгоритм сортировки,
  '  в реальных проектах может потребоваться более эффективный)
  For i = 1 To folderCount - 1
    For j = i + 1 To folderCount
      If fso.GetFolder(FolderArray(i)).DateCreated < fso.GetFolder(FolderArray(j)).DateCreated Then
        Temp = FolderArray(i)
        FolderArray(i) = FolderArray(j)
        FolderArray(j) = Temp
      End If
    Next j
  Next i

  ' Перемещаю указанное количество подпапок
  For i = 1 To WorksheetFunction.Min(NumFoldersToMove, folderCount) ' Используем Min, чтобы избежать ошибок, если подпапок меньше, чем запрошено
    fso.MoveFolder FolderArray(i), DestinationPath & "\" & fso.GetFolder(FolderArray(i)).Name
  Next i
  
 ' После перемещения можно открыть папку конечную.
  If MsgBox("Перемещение подпапок завершено! Открыть конечную папку?", vbYesNo, "Имя приложения: " & Application.Name) = vbNo Then Exit Sub
     shell "explorer.exe """ & DestinationPath & """", vbNormalFocus
  

  ' Освобождаю объекты
  Set fso = Nothing
  Set SourceFolder = Nothing
  Set subfolder = Nothing
  Set Subfolders = Nothing
End Sub


'   Purpose:
'   Копирование указанного количества папок с файлами из исходного каталога в целевой с добавлением к имени файлов цифры, если имена одинаковые.
Sub CopyFoldersWithFiles(control As IRibbonControl)
'   Sub CopyFoldersWithFiles()
    Dim SourcePath As String
    Dim destPath As String
    Dim fso As Object
    Dim SourceFolder As Object
    Dim destFolder As Object
    Dim file As Object
    Dim subfolder As Object
    Dim newFileName As String
    Dim fileExists As Boolean
    Dim counter As Integer
    Dim folderCount As Long
    Dim i As Long
   
    SourcePath = "C:\Users\s.lazarev\Desktop\ПОИСК первички\СКАНЫ для подготовки к сливу"  ' исходный
        Dim ИмяКонечнойПапки As String, ИмяПапкиСегодня As String
            ИмяКонечнойПапки = "Q:\LP2\задача 51677\НСК\"
            ИмяПапкиСегодня = Format(Now, "dd.MM.yyyy")
    destPath = "Q:\LP2\задача 51677\НСК\" & ИмяПапкиСегодня  ' конечный
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Проверка на существование исходной папки
    If Not fso.FolderExists(SourcePath) Then
        MsgBox "Исходная папка не существует."
        Exit Sub
    End If
    
    ' Создам целевую папку, если её нет
    If Not fso.FolderExists(destPath) Then
        fso.CreateFolder destPath
    End If
    
    ' Открою сразу папку
    shell "explorer.exe """ & destPath & """", vbNormalFocus
    
    Set SourceFolder = fso.GetFolder(SourcePath)
     If MsgBox("Стартую?", vbYesNo, "Имя приложения: " & Application.Name) = vbNo Then Exit Sub
    ' Запрос количества папок для копирования
    folderCount = Application.InputBox("Введите количество папок для копирования:", "Количество папок", , Type:=1)
    
    ' Проверяю, что введенное значение положительное
    If folderCount <= 0 Then
        MsgBox "Скоко папок копировать? (Положительное число, иначе - exit)."
        Exit Sub
    End If

    i = 0

    ' Копирую папки и файлы
    For Each subfolder In SourceFolder.Subfolders
        If i >= folderCount Then Exit For
        
        ' Проверяю, существует ли папка в целевом каталоге
        If Not fso.FolderExists(destPath & "\" & subfolder.Name) Then
            ' Если нет, создаю ее
            Set destFolder = fso.CreateFolder(destPath & "\" & subfolder.Name)
        Else
            ' Если да, использую существующую папку
            Set destFolder = fso.GetFolder(destPath & "\" & subfolder.Name)
        End If
        
        ' Копирую файлы
        For Each file In subfolder.Files
            newFileName = file.Name
            fileExists = True
            counter = 1
            
            ' Проверка на существование файла с таким именем
            Do While fileExists
                If fso.fileExists(destFolder.Path & "\" & newFileName) Then
                    ' Если файл существует, добавляю номер к имени
                    newFileName = Left(file.Name, InStrRev(file.Name, ".") - 1) & "_" & counter & Mid(file.Name, InStrRev(file.Name, "."))
                    counter = counter + 1
                Else
                    fileExists = False
                End If
            Loop
            
            ' Копирую файл с новым именем
            file.Copy destFolder.Path & "\" & newFileName
        Next file
        
        i = i + 1 ' Увеличиваю счетчик
    Next subfolder
    
    MsgBox "Копирование завершено. Скопировано " & i & " папок."
End Sub

'   Purpose:  Q:\LP2\Результаты сверки портфелей с августа 2020\Газпром ГПБ 2
'   Копирование указанного количества папок с файлами из исходного каталога в целевой с добавлением к имени файлов цифры, если имена одинаковые.
Sub КопированиеСливПапокСФайламиГПБ(control As IRibbonControl)
'   Sub CopyFoldersWithFilesГПБ() 'Копирование папок с файлами
    Dim SourcePath As String
    Dim destPath As String
    Dim fso As Object
    Dim SourceFolder As Object
    Dim destFolder As Object
    Dim file As Object
    Dim subfolder As Object
    Dim newFileName As String
    Dim fileExists As Boolean
    Dim counter As Integer
    Dim folderCount As Long
    Dim i As Long
    Dim renamedFilesCount As Long ' Добавлено: Счетчик переименованных файлов
    If MsgBox("Целевая папка Q:\LP2\Результаты сверки портфелей с августа 2020\Газпром ГПБ 2\сканы 10.09 Верно?", vbYesNo, "Имя приложения: " & Application.Name) = vbNo Then Exit Sub
    SourcePath = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_за_вчера"  ' исходный
''        Dim ИмяКонечнойПапки As String, ИмяПапкиСегодня As String
''            ИмяКонечнойПапки = "Q:\LP2\Результаты сверки портфелей с августа 2020\Газпром ГПБ 2\"
''            ИмяПапкиСегодня = Format(Now, "dd.MM")  ' ИмяПапкиСегодня = Format(Now, "dd.MM.yyyy")
''    destPath = "Q:\LP2\Результаты сверки портфелей с августа 2020\Газпром ГПБ 2\" & "сканы " & ИмяПапкиСегодня  ' конечный
    destPath = "Q:\LP2\Результаты сверки портфелей с августа 2020\Газпром ГПБ 2\сканы 10.09" ' конечный
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(SourcePath) Then  ' Проверка на существование исходной папки
        MsgBox "Исходная папка не существует."
        Exit Sub
    End If
    
    If Not fso.FolderExists(destPath) Then ' Создам целевую папку, если её нет
        fso.CreateFolder destPath
    End If

    Set SourceFolder = fso.GetFolder(SourcePath)
     If MsgBox("Стартую? (Копирование указанного количества папок с файлами из исходного каталога в целевой с добавлением к имени файлов цифры, если имена одинаковые.)", vbYesNo, "Имя приложения: " & Application.Name) = vbNo Then Exit Sub
    
    ' Запрос количества папок для копирования
    folderCount = Application.InputBox("Введите количество папок для копирования:", "Количество папок", , Type:=1)
    
    ' Проверяю, что введенное значение положительное
    If folderCount <= 0 Then
        MsgBox "Скоко папок копировать? (Положительное число, иначе - exit)."
        Exit Sub
    End If
    
    shell "explorer.exe """ & destPath & """", vbNormalFocus ' Открою сразу папку
    
    i = 0
    renamedFilesCount = 0 ' Инициализация счетчика
    
    ' Копирую папки и файлы
    For Each subfolder In SourceFolder.Subfolders
        If i >= folderCount Then Exit For
        
        ' Проверяю, существует ли папка в целевом каталоге
        If Not fso.FolderExists(destPath & "\" & subfolder.Name) Then
            ' Если нет, создаю ее
            Set destFolder = fso.CreateFolder(destPath & "\" & subfolder.Name)
        Else
            ' Если да, использую существующую папку
            Set destFolder = fso.GetFolder(destPath & "\" & subfolder.Name)
        End If
        
        ' Копирую файлы
        For Each file In subfolder.Files
            newFileName = file.Name
            fileExists = True
            counter = 1
            
            ' Проверка на существование файла с таким именем
            Do While fileExists
                If fso.fileExists(destFolder.Path & "\" & newFileName) Then
                    ' Если файл существует, добавляю номер к имени
                    newFileName = Left(file.Name, InStrRev(file.Name, ".") - 1) & "_" & counter & Mid(file.Name, InStrRev(file.Name, "."))
                    counter = counter + 1
                    renamedFilesCount = renamedFilesCount + 1 ' Увеличиваем счетчик переименованных
                Else
                    fileExists = False
                End If
            Loop
            
            ' Копирую файл с новым именем
            file.Copy destFolder.Path & "\" & newFileName
        Next file
        
        i = i + 1 ' Увеличиваю счетчик
    Next subfolder
    
    MsgBox "Копирование завершено. Скопировано " & i & " папок. Переименовано " & renamedFilesCount & " файлов.", vbYesNo, "© Alles gemacht!"
End Sub


'Option Explicit
 
Sub ПереместитьНесколькоПапокПоДате(control As IRibbonControl)
'Sub MoveFoldersByDate()
'Purpose - переместить определенное количество папок из исходного каталога (каталог А) в целевой каталог (каталог Б),
'основываясь на заданном критерии. В данном случае буду использовать дату создания папки в качестве критерия сортировки и перемещения.
 
    Dim fso As Object, SourceFolder As Object, TargetFolder As Object
    Dim FolderCollection As Object, folder As Object
    Dim SourcePath As String, TargetPath As String
    Dim NumFoldersToMove As Variant, i As Long, j As Long
    Dim FolderArray() As Variant, TempFolder As Variant
    
    ' Создаю объект FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Получаю исходный каталог от пользователя
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Выберите исходный каталог"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SourcePath = .SelectedItems(1)
        Else
            MsgBox "Перемещение отменено.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Получаю целевой каталог от пользователя
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Выберите целевой каталог"
        .AllowMultiSelect = False
        If .Show = -1 Then
            TargetPath = .SelectedItems(1)
        Else
            MsgBox "Перемещение отменено.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Получаю количество папок для перемещения от пользователя
    NumFoldersToMove = InputBox("Введите количество папок для перемещения:", "Количество папок", 1)
    If NumFoldersToMove = "" Then
        MsgBox "Перемещение отменено.", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(NumFoldersToMove) Then
        MsgBox "Некорректный ввод. Введите число.", vbExclamation
        Exit Sub
    End If
    NumFoldersToMove = CLng(NumFoldersToMove)
    
    ' Устанавливаю объекты папок
    Set SourceFolder = fso.GetFolder(SourcePath)
    Set TargetFolder = fso.GetFolder(TargetPath)
    
    ' Получаю коллекцию папок в исходном каталоге
    Set FolderCollection = SourceFolder.Subfolders
    
    ' Проверяю, есть ли папки в исходном каталоге
    If FolderCollection.count = 0 Then
        MsgBox "В исходном каталоге нет папок.", vbExclamation
        Exit Sub
    End If
    
    ' Создаю массив для хранения папок и их дат создания
    ReDim FolderArray(1 To FolderCollection.count, 1 To 2)
    
    ' Заполняю массив данными о папках и датах их создания
    i = 1
    For Each folder In FolderCollection
        FolderArray(i, 1) = folder.Name
        FolderArray(i, 2) = folder.DateCreated
    i = i + 1
    Next folder
    
    ' Сортирую массив по дате создания (от старых к новым)
    For i = 1 To UBound(FolderArray, 1) - 1
    For j = i + 1 To UBound(FolderArray, 1)
        If FolderArray(i, 2) > FolderArray(j, 2) Then
            ' Меняю местами папки
            TempFolder = FolderArray(i, 1)
            FolderArray(i, 1) = FolderArray(j, 1)
            FolderArray(j, 1) = TempFolder
            
            ' Меняю местами даты
            TempFolder = FolderArray(i, 2)
            FolderArray(i, 2) = FolderArray(j, 2)
            FolderArray(j, 2) = TempFolder
        End If
    Next j
    Next i
    
    ' Перюещаю заданное количество папок
    On Error Resume Next ' Обработка ошибок перемещения (например, если папка уже существует)
    For i = 1 To WorksheetFunction.Min(NumFoldersToMove, FolderCollection.count) ' Использую Min, чтобы не выйти за границы массива
    Set folder = fso.GetFolder(SourcePath & "\" & FolderArray(i, 1)) ' Получаю объект папки
    fso.MoveFolder SourcePath & "\" & FolderArray(i, 1), TargetPath & "\" & FolderArray(i, 1) ' Перюещаю папку
        If Err.Number <> 0 Then
'        Debug.Print "Ошибка при перемещении папки: " & FolderArray(i, 1) & " - " & Err.Description
        Err.Clear
        End If
    Next i
    On Error GoTo 0 ' Отключаю обработку ошибок
    
    MsgBox "Перемещение завершено.", vbInformation
    
    ' Очищаю объекты
    Set fso = Nothing
    Set SourceFolder = Nothing
    Set TargetFolder = Nothing
    Set FolderCollection = Nothing
    Set folder = Nothing
 
End Sub



Sub ПереместитьПустыеИлиZeroFileПодпапки(control As IRibbonControl)     ' ПОИСК ОТЛИЧКА
'макрос , который перемещает подпапки, либо имеющие размер 0 байт,
'либо содержащие внутри себя хотя бы один файл с размером 0 байт, из одной папки в другую:

If MsgBox("перемещает подпапки, либо имеющие размер 0 байт, либо содержащие внутри себя хотя бы один файл с размером 0 байт, из одной папки в другую, продолжить?", vbYesNo) = vbNo Then Exit Sub

  Dim fso As Object, SourceFolder As Object, destFolder As Object, subfolder As Object
  Dim file As Object
  Dim SourcePath As String, destPath As String
  Dim HasZeroFile As Boolean

  SourcePath = "C:\Users\s.lazarev\Desktop\ПОИСК первички\СКАНЫ для подготовки к сливу"  ' Путь к папке, откуда перемещаем
  destPath = "C:\Users\s.lazarev\Desktop\ПОИСК первички\СКАНЫ для вырезки\04.12.2025"    ' Путь к папке, куда перемещаем

  ' Создаем объекты FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Проверяем, существуют ли исходная и целевая папки
  If Not fso.FolderExists(SourcePath) Then
    MsgBox "Исходная папка не найдена: " & SourcePath, vbCritical
    Exit Sub
  End If

  If Not fso.FolderExists(destPath) Then
    MsgBox "Целевая папка не найдена: " & destPath, vbCritical
    Exit Sub
  End If

  ' Получаем объекты папок
  Set SourceFolder = fso.GetFolder(SourcePath)
  Set destFolder = fso.GetFolder(destPath)

  ' Перебираем все подпапки в исходной папке
  For Each subfolder In SourceFolder.Subfolders
    HasZeroFile = False ' Сбрасываем флаг для каждой подпапки

    ' Проверяем, есть ли в папке файлы с нулевым размером
    For Each file In subfolder.Files
      If file.Size = 0 Then
        HasZeroFile = True
        Exit For ' Выходим из цикла, если нашли хотя бы один файл с нулевым размером
      End If
    Next file

    ' Проверяем размер папки (0 байт) ИЛИ наличие файла с 0 байт
    If subfolder.Size = 0 Or HasZeroFile Then
      ' Перемещаем папку
      On Error Resume Next ' Обработка ошибок, если папка уже существует в целевой папке
      fso.MoveFolder subfolder.Path, destPath & "\" & subfolder.Name
      On Error GoTo 0 ' Отключаем обработку ошибок

''      If Err.Number = 0 Then     ' ошибки
''        Debug.Print "Перемещена папка: " & SubFolder.Name
''      Else
''        Debug.Print "Ошибка при перемещении папки: " & SubFolder.Name & " - " & Err.Description
''        Err.Clear
''      End If
    End If
  Next subfolder

  MsgBox "Перемещены подпапки, имеющие размер 0 байт, либо содержащие внутри себя хотя бы один файл с размером 0 байт!", vbInformation

  ' Очищаем объекты
  Set fso = Nothing
  Set SourceFolder = Nothing
  Set destFolder = Nothing
  Set subfolder = Nothing
  Set file = Nothing

End Sub


Sub FileInfo() 'Получение свойств любого файла программным способом на примере файла "Схема Белоусова.png", расположенного в папке "C:\Users\Evgeniy\Downloads\":
Dim ns As Object, i As Integer, n As Integer
Set ns = CreateObject("Shell.Application").Namespace("C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\")
    For i = 0 To 303
        If ns.GetDetailsOf(ns.ParseName("ЛАЗАРЕВ_VII_03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.XLSX"), i) <> "" Then
            n = n + 1
            Cells(n, 1) = ns.GetDetailsOf("ЛАЗАРЕВ_VII_03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.XLSX", i) & " = " & ns.GetDetailsOf(ns.ParseName("ЛАЗАРЕВ_VII_03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.XLSX"), i)
        End If
    Next
Set ns = Nothing
End Sub


Sub УдалитьПустыеПапкиУказавКаталог()   ' РАБОТАЕТ ВЕРНО

Dim fso As Object, objFolder As Object, objSubFolder As Object
Dim ПутьККаталогу As String
Dim КоличествоУдаленныхПапок As Long
Dim objShell As Object, objFolderBrowse As Object
Dim DefaultPath As String

' Создаем объект FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Указываем папку по умолчанию
DefaultPath = "C:\Users\s.lazarev\Desktop\ПОИСК первички\"  ' НЕ РАБОТАЕТ !!!

' Создаем объект Shell.Application для открытия диалогового окна выбора папки
Set objShell = CreateObject("Shell.Application")
Set objFolderBrowse = objShell.BrowseForFolder(0, "Выберите каталог для удаления пустых папок:", 0, DefaultPath)

' Проверяем, выбрал ли пользователь каталог
If Not objFolderBrowse Is Nothing Then
ПутьККаталогу = objFolderBrowse.Self.Path
Else
MsgBox "Действие отменено. Каталог не выбран.", vbExclamation
Exit Sub
End If

' Проверяем, существует ли указанный каталог
If Not fso.FolderExists(ПутьККаталогу) Then
MsgBox "Указанный каталог не существует.", vbExclamation
Exit Sub
End If

' Инициализируем счетчик удаленных папок
КоличествоУдаленныхПапок = 0

' Вызываем рекурсивную функцию для удаления пустых папок
УдалитьПустыеПапкиРекурсивно ПутьККаталогу, fso, КоличествоУдаленныхПапок

' Выводим сообщение о количестве удаленных папок
MsgBox "Удалено " & КоличествоУдаленныхПапок & " пустых папок.", vbInformation

' Очищаем объекты
Set objFolder = Nothing
Set objSubFolder = Nothing
Set fso = Nothing
Set objShell = Nothing
Set objFolderBrowse = Nothing

End Sub

' Рекурсивная функция для удаления пустых папок
Private Sub УдалитьПустыеПапкиРекурсивно(Путь As String, fso As Object, ByRef КоличествоУдаленныхПапок As Long)
Dim objFolder As Object, objSubFolder As Object

Set objFolder = fso.GetFolder(Путь)

' Перебираем подпапки в текущей папке (в обратном порядке, чтобы не нарушить коллекцию при удалении)
For Each objSubFolder In objFolder.Subfolders
' Рекурсивно вызываем функцию для подпапки
УдалитьПустыеПапкиРекурсивно objSubFolder.Path, fso, КоличествоУдаленныхПапок
Next objSubFolder

' После обработки подпапок, проверяем, пуста ли текущая папка
If objFolder.Subfolders.count = 0 And objFolder.Files.count = 0 Then
' Удаляем пустую папку
On Error Resume Next ' Игнорируем ошибки (например, если нет прав на удаление)
fso.DeleteFolder objFolder.Path
On Error GoTo 0 ' Возвращаем обработку ошибок

' Проверяем, была ли удалена папка
If Err.Number = 0 Then
КоличествоУдаленныхПапок = КоличествоУдаленныхПапок + 1
End If
End If

Set objFolder = Nothing
Set objSubFolder = Nothing

End Sub

Sub СчетПустыхПапокВКаталогеПоВыбору() 'CountEmptyFolders()  +  ' Функция проверки что папка пустая:  Function IsFolderEmpty(FolderPath As String) As Boolean
    Dim oFD As FileDialog
    Dim x, lf As Long
    
    Dim fso As Object, folder As Object, subfolder As Object
  Dim FolderPath As String, EmptyFolderCount As Long
  Dim DefaultPath As String

    'назначаем переменной ссылку на экземпляр диалога
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    With oFD 'используем короткое обращение к объекту
    'так же можно без oFD
    'With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Выбрать папку с отчетами" '"заголовок окна диалога
        .ButtonName = "Выбрать папку"
        .Filters.Clear 'очищаем установленные ранее типы файлов
        .InitialFileName = "C:\Users\s.lazarev\Desktop" '"назначаем первую папку отображения
        .InitialView = msoFileDialogViewLargeIcons 'вид диалогового окна(доступно 9 вариантов)
        If .Show = 0 Then Exit Sub 'показывает диалог
        'цикл по коллекции выбранных в диалоге файлов
        FolderPath = .SelectedItems(1) 'считываем путь к папке
        MsgBox "Выбрана папка: '" & FolderPath & "'", vbInformation, "www.excel-vba.ru"
    End With
      ' Инициализируем счетчик пустых папок
  EmptyFolderCount = 0
  ' Создаем объект FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Рекурсивно обходим все папки и подпапки
  For Each folder In fso.GetFolder(FolderPath).Subfolders
    If ЯвляетсяЛиПапкаПустой(folder.Path) Then
      EmptyFolderCount = EmptyFolderCount + 1
    End If
  Next folder

  ' Выводим результат в сообщении
  MsgBox "Количество пустых папок в папке '" & FolderPath & "': " & EmptyFolderCount, vbInformation

  ' Освобождаем объекты
  Set fso = Nothing
  Set folder = Nothing
  Set subfolder = Nothing

End Sub

Sub CountMacrosInProject() 'Макрос для подсчёта макросов в проекте....................... Макрос для подсчёта макросов в проекте
    Dim VBProj As Object
    Dim VBComp As Object
    Dim CodeMod As Object
    Dim LineNum As Long
    Dim ProcName As String
    Dim ProcKind As Long
    Dim MacroCount As Long
    
    ' Получаем текущий проект VBA
    Set VBProj = ThisWorkbook.VBProject
    
    ' Проверяем, разрешён доступ к объектной модели VBA
    If Not VBProj Is Nothing Then
        MacroCount = 0
        
        ' Проходим по всем компонентам проекта (модули, листы, книга и т. д.)
        For Each VBComp In VBProj.VBComponents
            ' Получаем модуль кода компонента
            Set CodeMod = VBComp.CodeModule
            
            ' Начинаем с первой строки кода
            LineNum = 1
            
            ' Ищем все процедуры и функции в модуле
            Do Until LineNum >= CodeMod.CountOfLines
                ProcName = CodeMod.ProcOfLine(LineNum, ProcKind)
                
                ' Если найдена процедура или функция
                If ProcName <> "" Then
                    ' Увеличиваем счётчик
                    MacroCount = MacroCount + 1
                    
                    ' Переходим к следующей строке после текущей процедуры
                    LineNum = CodeMod.ProcStartLine(ProcName, ProcKind) + CodeMod.ProcCountLines(ProcName, ProcKind)
                Else
                    ' Иначе переходим к следующей строке
                    LineNum = LineNum + 1
                End If
            Loop
        Next VBComp
        
        ' Выводим результат
        MsgBox "Количество макросов (Sub и Function) в проекте ""Надстройка2"": " & MacroCount, vbInformation, "Результат подсчёта"
    Else
        MsgBox "Доступ к VBProject запрещён. Включите «Доверие к объектной модели проекта VBA» в настройках Excel.", vbExclamation, "Ошибка"
    End If
End Sub





' Функция для проверки, является ли папка пустой  IsFolderEmpty
Function ЯвляетсяЛиПапкаПустой(FolderPath As String) As Boolean
  Dim fso As Object, folder As Object

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set folder = fso.GetFolder(FolderPath)

  ЯвляетсяЛиПапкаПустой = (folder.Files.count = 0 And folder.Subfolders.count = 0)

  Set fso = Nothing
  Set folder = Nothing
End Function
