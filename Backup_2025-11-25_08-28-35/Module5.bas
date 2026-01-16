Attribute VB_Name = "Module5"
Public Sub ZZZввв() ' отсюда буду удалять строку
    Exit Sub
End Sub

Public Sub ZZZ() ' отсюда буду удалять строку
      Dim helpFilePath As String
    On Error GoTo ErrHandler
     If InputBox("Введите пароль Администратора", "Аторизация") <> "123" Then
     Const lSeconds As Long = 3
                MessageBoxTimeOut 0, "Не правильный пароль," & _
                vbNewLine & "в доступе отказано!" & _
                vbNewLine & " " & _
                vbNewLine & "Это окно закроется через 3 секунды.", "Проверка прав доступа", _
                vbInformation + vbOKOnly, 0&, lSeconds * 1000: Exit Sub
    End If
        helpFilePath = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns" & "\2.chm"
        HtmlHelp 0, helpFilePath, 0, 0
         ' остальной код
    Exit Sub
ErrHandler:     MsgBox "ТЕКСТ"
     MsgBox "Ошибка: " & Err.Description & _
     vbNewLine & "Имя Модуля: " & ИмяМодуля & _
     vbNewLine & "Имя Процедуры: " & ИмяПроцедуры, vbExclamation, "Ошибка выполнения"
    
    On Error GoTo 0  ' Сброс обработчика ошибок
    Exit Sub
End Sub

Public Sub ZZ() ' отсюда буду удалять строку
   Dim helpFilePath As String
   On Error GoTo ErrHandler  ' буду удалять строку
     If InputBox("Введите пароль Администратора", "Аторизация") <> "123" Then
     Const lSeconds As Long = 3
                MessageBoxTimeOut 0, "Не правильный пароль," & _
                vbNewLine & "в доступе отказано!" & _
                vbNewLine & " " & _
                vbNewLine & "Это окно закроется через 3 секунды.", "Проверка прав доступа", _
                vbInformation + vbOKOnly, 0&, lSeconds * 1000: Exit Sub
    End If
        helpFilePath = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns" & "\2.chm"
        HtmlHelp 0, helpFilePath, 0, 0
         ' остальной код
    Exit Sub
ErrHandler:     MsgBox "ТЕКСТ"
     MsgBox "Ошибка: " & Err.Description & _
     vbNewLine & "Имя Модуля: " & ИмяМодуля & _
     vbNewLine & "Имя Процедуры: " & ИмяПроцедуры, vbExclamation, "Ошибка выполнения"
    
    On Error GoTo 0  ' Сброс обработчика ошибок
    Exit Sub
End Sub



