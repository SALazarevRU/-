Attribute VB_Name = "m9_Проверка_UserName_Пароль"
Sub Пароль()
    If InputBox("Введите пароль Администратора") <> "123" Then MsgBox "Неправильный пароль": Exit Sub
End Sub

Sub ПроверкаИмениПользователя_CheckUserName()
    
    Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev6"
    username = Environ("UserName")  ' Получаем имя пользователя.
        If username = SpecifiedUserName Then
              ' Если имя совпадает, выполнить дальнейший код.
        Else  ' Иначе:
'                MsgBox "Имя пользователя компьютера не совпадает с указанным", vbCritical
'                Const lSeconds As Long = 5
'                MessageBoxTimeOut 0, "Имя пользователя компьютера не совпадает с указанным!" & _
'                vbNewLine & "Программа будет остановлена." & _
'                vbNewLine & "Это окно закроется автоматически через 5 секунд.", "Сообщение", _
'                vbInformation + vbOKOnly, 0&, lSeconds * 1000
                CreateObject("WScript.Shell").Popup ("Нет доступа." & _
                vbNewLine & "Это окно закроется через 1 секунду"), 1, "Доступ заблокирован", 48
            Exit Sub  ' Остановка кода.
        End If
End Sub
