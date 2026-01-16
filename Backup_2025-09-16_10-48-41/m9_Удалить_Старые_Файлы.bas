Attribute VB_Name = "m9_Удалить_Старые_Файлы"
'Чтобы установить объект FileSystemObject (FSO) в VBA, можно использовать метод CreateObject:
'Перед работой с FSO в VBA необходимо **добавить ссылку на его библиотеку** в редакторе Visual Basic: выберите «Microsoft Scripting Runtime».


Sub УдалитьСтарыеФайлыБэкап(control As IRibbonControl)
    Dim folderPath As String
    Dim FileName As String
    Dim fileDate As Date
    Dim FSO As Object
    Dim file As Object
    
    ' Укажите путь к папке
'    folderPath = "C:\Users\Хозяин\Desktop\Бэкапы\" ' Измените на нужный путь
     folderPath = "C:\Users\s.lazarev\Desktop\Бэкапы\" ' Измените на нужный путь
    
    ' Создаем объект FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Проверяем, существует ли папка
    If FSO.FolderExists(folderPath) Then
        ' Проходим по всем файлам в папке
        FileName = Dir(folderPath & "*.*")
        Do While FileName <> ""
            Set file = FSO.GetFile(folderPath & FileName)
            fileDate = file.DateCreated
            
            ' Проверяем, создан ли файл более 30 минут назад
            If Now - fileDate > TimeValue("99:10:00") Then
                ' Удаляем файл
                file.Delete
            End If
            
            FileName = Dir ' Переходим к следующему файлу
        Loop
    Else
        MsgBox "Папка не найдена: " & folderPath
    End If
    
    ' Освобождаем объект
    Set FSO = Nothing
    
            Const lSeconds As Long = 3
            MessageBoxTimeOut 0, "Удалены старые файлы из папки:" & _
            vbNewLine & folderPath & _
            vbNewLine & " " & _
            vbNewLine & "Это окно закроется автоматически через 3 секунды", "Сообщение", _
            vbInformation + vbOKOnly, 0&, lSeconds * 1000
End Sub


Sub УдалитьСтарыеФайлыПанели(control As IRibbonControl)
    Dim folderPath As String
    Dim FileName As String
    Dim fileDate As Date
    Dim FSO As Object
    Dim file As Object
    
    ' Укажите путь к папке
    folderPath = "D:\OneDrive\ECXELнаOneDrive\Надстройки\Надстройка2\Бэкапы\" ' Измените на нужный путь
    
    ' Создаем объект FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Проверяем, существует ли папка
    If FSO.FolderExists(folderPath) Then
        ' Проходим по всем файлам в папке
        FileName = Dir(folderPath & "*.*")
        Do While FileName <> ""
            Set file = FSO.GetFile(folderPath & FileName)
            fileDate = file.DateCreated
            
            ' Проверяем, создан ли файл более 30 минут назад
            If Now - fileDate > TimeValue("00:23:00") Then
                ' Удаляем файл
                file.Delete
            End If
            
            FileName = Dir ' Переходим к следующему файлу
        Loop
    Else
        MsgBox "Папка не найдена: " & folderPath
    End If
    
    ' Освобождаем объект
    Set FSO = Nothing
    
            Const lSeconds As Long = 3
            MessageBoxTimeOut 0, "Удалены старые файлы из папки:" & _
            vbNewLine & folderPath & _
            vbNewLine & " " & _
            vbNewLine & "Это окно закроется автоматически через 3 секунды", "Сообщение", _
            vbInformation + vbOKOnly, 0&, lSeconds * 1000
End Sub

Sub Удалить_файлы_в_папке_ТЕМП_по_маске(control As IRibbonControl)  '   удалить файлы в папке TEMP по маске, очистить папку TEMP ТЕМП по маске .jpg
    Dim counter
    Dim fn
     If MsgBox("Удалить файлы в папке TEMP по маске?", vbYesNo) = vbNo Then Exit Sub
    ChDir ("C:\Users\s.lazarev\AppData\Local\Temp\")
         
    fn = Dir("*.JPG")
    counter = 0
     
    While Len(fn) > 0
     
    counter = counter + 1
    fn = Dir()
    Wend
    
    '    удалить файлы в папке TEMP по маске:
    If MsgBox("Общее количество файлов .JPG в папке TEMP :    " & counter & "                 " & vbNewLine & "Удалить их?", vbYesNo) = vbNo Then Exit Sub
    shell "cmd /c del """ & Chr(34) & "%TEMP%\*.jpg" & Chr(34)

'    очистить Корзину
    If MsgBox("Очистить корзину от файлов в ней находящихся?", vbYesNo) = vbNo Then Exit Sub
    SHEmptyRecycleBin 0, "C:\", 1
End Sub

Sub УдалитьСтарыеФайлыTemp(control As IRibbonControl) ' Удалить Старые Файлы в папке Temp.
    Dim folderPath As String
    Dim FileName As String
    Dim fileDate As Date
    Dim FSO As Object
    Dim file As Object
    On Error Resume Next
'    On Error GoTo ErrHandler
    
    ' Укажите путь к папке
    folderPath = "C:\Users\s.lazarev\AppData\Local\Temp\" ' Измените на нужный путь
    
    ' Создаем объект FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Проверяем, существует ли папка
    If FSO.FolderExists(folderPath) Then
        ' Проходим по всем файлам в папке
        FileName = Dir(folderPath & "*.*")
        Do While FileName <> ""
            Set file = FSO.GetFile(folderPath & FileName)
            fileDate = file.DateCreated
            
            ' Проверяем, создан ли файл более 7 ч 00 минут назад
            If Now - fileDate > TimeValue("03:00:00") Then
                ' Удаляем файл
                file.Delete
            End If
            
            FileName = Dir ' Переходим к следующему файлу
        Loop
    Else
        MsgBox "Папка не найдена: " & folderPath
    End If
    
    ' Освобождаем объект
    Set FSO = Nothing
    Const lSeconds As Long = 3
            MessageBoxTimeOut 0, "Удалены старые файлы из папки:" & _
            vbNewLine & folderPath & _
            vbNewLine & " " & _
            vbNewLine & "Это окно закроется автоматически через 3 секунды", "Сообщение", _
            vbInformation + vbOKOnly, 0&, lSeconds * 1000
'    Exit Sub
'ErrHandler:
'    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Ошибка выполнения"
'   On Error GoTo 0
'   Exit Sub
    
            
End Sub

Sub Delete_Temp_Files_Primitive()

Dim sFileType As String ' объявить тип файла
Dim sTempDir As String ' временный каталог

'On Error Resume Next

sFileType = "*.tmp"
sTempDir = "C:\Users\s.lazarev\AppData\Local\Temp\"

: Kill sTempDir & sFileType

End Sub


' Private Declare Function SHEmptyRecycleBinW Lib "shell32.dll" (ByVal hwnd As Long, ByVal Path As String, ByVal Flags As Long) As Long
'  корзина,

Sub ClearBasket() ' макрос, который очистит корзину----------РАБОТАЕТ
SHEmptyRecycleBin 0, "C:\", 1
'SHEmptyRecycleBinW 0, "", 3
End Sub
'где hwnd - дескриптор окна (используйте 0 или дескриптор своего окна);
'Path - диск на котором будет очищаться корзина;
'Flags -флаги.
'Если Вам нужно обычная очистка "корзины", т.е. с окном подтверждения очистки, то вставьте
'SHEmptyRecycleBinW 0, "", 0
'Если не нужно показывать окно подтверждения, то вставьте
'SHEmptyRecycleBinW 0, "", 1
'Если Вам не нужно чтобы пользователь видел процесс удаления файлов из "корзины", вставьте
'SHEmptyRecycleBinW 0, "", 2
'Чтобы не было звука "Очистка корзины", вставьте такой код:
'SHEmptyRecycleBinW 0, "", 4


Sub ОчиститьКорзину()
    Const SHERB_NOCONFIRMATION = &H1
    Const SHERB_NOPROGRESSUI = &H2
    Const SHERB_NOSOUND = &H4
'    Комбинация флагов:
    
'    SHERB_NOCONFIRMATION ' Не показывать подверждение удаления файлов пользователю
'    SHERB_NOPROGRESSUI ' Не показывать диалоговое окно,показывающее процесс удаления файлов из Корзины
'    SHERB_NOSOUND ' Не воспроизводить звук после удаления файлов из Корзины
'    Код очистки Корзины выглядит так:
    
    a = SHEmptyRecycleBin(0, "", SHERB_NOPROGRESSUI)
End Sub


