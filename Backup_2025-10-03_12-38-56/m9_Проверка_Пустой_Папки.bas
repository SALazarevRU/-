Attribute VB_Name = "m9_Проверка_Пустой_Папки"
'Пример использования Exit Sub для проверки пустой папки:

Sub Button1_click()
   Dim FileSystem As Object
   Set FileSystem = CreateObject("Scripting.FileSystemObject")
   strFolderPath = Application.ActiveWorkbook.Path
   Set oFolder = FileSystem.GetFolder(strFolderPath)
    If (oFolder.Subfolders.Count = 0) Then
      MsgBox "Folder is empty!", vbOKOnly + vbInformation, "Information!"
      Exit Sub
    End If
End Sub

'Пример проверки пустой папки:
Sub Checking_An_Empty_Folder() ' 1)Show Folder Dialog, 2)
    Dim oFD As FileDialog
    Dim x, lf As Long
    'назначаем переменной ссылку на экземпляр диалога
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    With oFD 'используем короткое обращение к объекту
    'так же можно без oFD
    'With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Выбрать папку с отчетами" '"заголовок окна диалога
        .ButtonName = "Выбрать папку"
        .Filters.Clear 'очищаем установленные ранее типы файлов
        .InitialFileName = "C:\Temp\" '"назначаем первую папку отображения
        .InitialView = msoFileDialogViewLargeIcons 'вид диалогового окна(доступно 9 вариантов)
        If .Show = 0 Then Exit Sub 'показывает диалог
        'цикл по коллекции выбранных в диалоге файлов
        x = .SelectedItems(1) 'считываем путь к папке
        MsgBox "Выбрана папка: '" & x & "'", vbInformation, "www.excel-vba.ru"
    End With
    
   Dim FileSystem As Object
   Set FileSystem = CreateObject("Scripting.FileSystemObject")
   strFolderPath = x  'Application.ActiveWorkbook.Path
   Set oFolder = FileSystem.GetFolder(strFolderPath)
    If (oFolder.Subfolders.Count = 0) Then
      MsgBox "Папка пуста!", vbOKOnly + vbInformation, "Information!"
      Exit Sub
    End If
    
End Sub
