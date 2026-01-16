Attribute VB_Name = "m9_Проверка_ЮзерНаме_список"
Option Explicit

Sub CheckName_1()
    Dim username As Variant
    Dim usernames  ' массив разрешённых имён пользователей
    Dim accessAllowed As Boolean  ' логическое значение, указывающее, разрешён ли доступ
    accessAllowed = False
    
    usernames = Array("s.lazarev", "Bob", "Dave", "Sally", "Amanda")  ' список разрешённых имён
    
    For Each username In usernames
        If Environ("USERNAME") = username Then
            accessAllowed = True
            Exit For
        End If
    Next
    
    If Not accessAllowed Then
    ' действие при отсутствии разрешения
         MsgBox "Имя пользователя не равно одному из разрешённых, Exit Sub"
    End If
     MsgBox "Имя пользователя равно одному из разрешённых, приветствую" ' Далее выполняется код, если имя пользователя входит в список
End Sub

Sub CheckName_2() 'проверка имени пользователя и выхода из процедуры, если оно не равно одному из разрешённых
    Dim allowedNames As Variant, username, anyOfAllowedNames As String ' Список разрешённых имён
'    Dim userName As String  ' создаем переменную userName со строковым типом данных
'    Dim anyOfAllowedNames As String 'тип переменной anyOfAllowedNames - String при использовании оператора Option Explicit.
    
    allowedNames = Array("s.lazarev", "Bob", "Dave", "Sally", "Amanda")  ' список разрешённых имён
    
    username = Environ("USERNAME")  ' Получаем имя пользователя.
    Environ ("USERNAME") ' Получаем имя пользователя.
'    Если имя равно одному из элементов массива, сообщение не появится
    If username <> anyOfAllowedNames Then ' anyOfAllowedNames функция в языке VBA, которая возвращает любое из допустимых имён.
        MsgBox "Имя пользователя не равно одному из разрешённых, Exit Sub"
        Exit Sub
    End If
     MsgBox "Имя пользователя равно одному из разрешённых, приветствую" ' Далее выполняется код, если имя пользователя входит в список
End Sub


Sub VarType_Example2()

    Dim MyVar

    Set MyVar = ThisWorkbook

    MsgBox VarType(MyVar)

End Sub
