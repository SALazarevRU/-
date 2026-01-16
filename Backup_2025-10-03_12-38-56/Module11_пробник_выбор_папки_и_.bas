Attribute VB_Name = "Module11_пробник_выбор_папки_и_"



Sub СчетПустыхПапокВКаталогеПоВыбору() 'CountEmptyFolders()
    Dim oFD As FileDialog
    Dim x, lf As Long
    
    Dim FSO As Object, Folder As Object, SubFolder As Object
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
  Set FSO = CreateObject("Scripting.FileSystemObject")

  ' Рекурсивно обходим все папки и подпапки
  For Each Folder In FSO.GetFolder(FolderPath).Subfolders
    If IsFolderEmpty(Folder.Path) Then
      EmptyFolderCount = EmptyFolderCount + 1
    End If
  Next Folder

  ' Выводим результат в сообщении
  MsgBox "Количество пустых папок в папке '" & FolderPath & "': " & EmptyFolderCount, vbInformation

  ' Освобождаем объекты
  Set FSO = Nothing
  Set Folder = Nothing
  Set SubFolder = Nothing

End Sub


' Функция для проверки, является ли папка пустой
Function IsFolderEmpty(FolderPath As String) As Boolean
  Dim FSO As Object, Folder As Object

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set Folder = FSO.GetFolder(FolderPath)

  IsFolderEmpty = (Folder.Files.Count = 0 And Folder.Subfolders.Count = 0)

  Set FSO = Nothing
  Set Folder = Nothing
End Function


