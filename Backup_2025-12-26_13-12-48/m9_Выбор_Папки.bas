Attribute VB_Name = "m9_Выбор_Папки"

Public PapkaSkanov, lf As String ' ОБЛАСТЬ ВИДИМОСТИ перменная становится доступной для использования в любом модуле и любой процедуре проекта, т.е. она глобальная

Sub ВыборПапкиСПапкамиСканов(control As IRibbonControl)
On Error Resume Next
    Dim oFD As FileDialog
    
    'назначаем переменной ссылку на экземпляр диалога
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    With oFD 'используем короткое обращение к объекту
    'так же можно без oFD
    'With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Выбрать папку с отчетами" '"заголовок окна диалога
        .ButtonName = "Выбрать папку"
        .Filters.Clear 'очищаем установленные ранее типы файлов
        .InitialFileName = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.06.2025\" '"назначаем первую папку отображения
        .InitialView = msoFileDialogViewLargeIcons 'вид диалогового окна(доступно 9 вариантов)
        If .Show = 0 Then Exit Sub 'показывает диалог
        'цикл по коллекции выбранных в диалоге файлов
        PapkaSkanov = .SelectedItems(1) 'считываем путь к папке
'        MsgBox "Выбрана папка: '" & PapkaSkanov & "'", vbInformation, "Сообщение"

    MsgBox "Выбрана папка " & PapkaSkanov
        
        
    End With
End Sub

Sub Выбор_ПАПКИ()
    Dim xFilePath As String
    Set xObjFD = Application.FileDialog(msoFileDialogFolderPicker)
    With xObjFD
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.count > 0 Then
            xFilePath = .SelectedItems.item(1)
            MsgBox "Выбрана папка: '" & xFilePath & "'", vbInformation, "Сообщение"
        Else
            MsgBox "Папка не выбрана"
            Exit Sub
        End If
    End With
End Sub


