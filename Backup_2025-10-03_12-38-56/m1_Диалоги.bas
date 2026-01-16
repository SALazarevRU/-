Attribute VB_Name = "m1_Диалоги"
'   ДИАЛОГИ. Тренинг.
' Popup ,2 ;
' MsgBox "  "
' AutoCloseMsgBox
' If MsgBox("Запустить?", vbYesNo) = vbNo Then Exit Sub
' If MsgBox("уверены?", vbYesNo, "У " & Application.Name & " к Вам пара вопросов: ") = vbNo Then Exit Sub
'У вас недостаточно прав для доступа к данному приложению. Обратитесь к администратору вашего Битрикс24.
'БЛОК ВЫБОРА:


Private Sub Вопр()
CreateObject("WScript.Shell").Popup "Это окно закроется через 5 секунд", 5, "Microsoft Excel", 48
CreateObject("WScript.Shell").Popup ("В настоящее время доступ заблокирован." & _
                vbNewLine & "Это окно сейчас закроется "), 1, "Информация о блокировке доступа к выполнению программы", 48
                
CreateObject("WScript.Shell").Popup "'Новый текстовый файл.txt' создан в папке Downloads!" & _
                    vbNewLine & _
                    vbNewLine & " Это окно закроется через 2 секунд", 2, "Сообщение о завершении операции", 64
End Sub


Private Sub Вопросы()

    If MsgBox("Когда сканирование завершится - нажмите ОК для переименования папки и скана в ней по имени клайма." _
            & vbNewLine & "Спасибо!", vbYesNo, "(.)(.) " & Application.Name & "  рекомендует Вам следующие действия: ") = vbNo Then Exit Sub

    Debug.Print "Имя приложения: "; Application.Name 'Листинг 13.1. Вывести имя приложения
    'Запрос. Нет - Exit Sub. Да - код выполниться далее со след строки
    
    If MsgBox("Запустить заполнение Динамики?", vbYesNo, "Имя приложения: " & Application.Name) = vbNo Then Exit Sub
    
    Call ЗаполнитьОтчётВФоне
End Sub

Sub ШАБЛОН()
   
End Sub

'VBA Получить местоположение текущей рабочей книги в Excel:
Sub Местоположение_этой_книги()
    
    'Variable declaration 'Объявление переменной:
    Dim sWorkbookLocation As String
    sWorkbookLocation = ThisWorkbook.FullName
    Debug.Print "Местоположение текущей рабочей книги в Excel : " & sWorkbookLocation, vbInformation, "VBAF1"

'   Проверка на случайное нажатие:
    If MsgBox("Вы уверены что хотите запустить поиск местонахождения этой книги?", vbYesNo, "У " & Application.Name & " к Вам пара вопросов: ") = vbNo Then Exit Sub
'   БЛОК ВЫБОРА:
        result = MsgBox("Да - смотреть путь к текущей книге" _
            & vbNewLine & "Нет - файл закроется", vbYesNoCancel, "Выберите, пожалуйста дальнейшие действия:")
        Select Case result
            Case vbYes
               MsgBox "Ок, полный путь к  книге: " & sWorkbookLocation ' показал путь
            Case vbNo
               MsgBox ("Вы выбрали НЕТ,  Exit Sub") ' .........
               Exit Sub
            Case vbCancel
                MsgBox ("Вы выбрали Отмена") ' .........
        End Select
    
'   Проверка на случайное нажатие
    If MsgBox("Запустить проверку наличия файла [Какая-то книга.xlsb]?", vbYesNo) = vbNo Then Exit Sub
   
'   Проверка наличия файла
    If Dir("C:\Users\Admin\Desktop\тестовый файл\Какая-то книга.xlsb") = "" Then '(Then-затем)
        MsgBox "Файл C:\Users\Admin\Desktop\тестовый файл\Значения.xlsb" & vbNewLine & "Файл не найден, Exit Sub", 48, "К нашему сожалению:"
        Exit Sub
    End If
End Sub
'=============================================================================================================================================
'
' Purpose   : Процедура вывода информации в msgBox с закрытием через 3 секунды
'                                         в окно Immediate Имя пользователя и компьютера
'---------------------------------------------------------------------------------------

'Declare PtrSafe Function MessageBoxTimeOut Lib "User32" Alias "MessageBoxTimeoutA" _
'                        (ByVal hwnd As Long, ByVal lpText As String, _
'                         ByVal lpCaption As String, ByVal uType As VbMsgBoxStyle, _
'                         ByVal wLanguageId As Long, ByVal dwMilliseconds As Long) As Long
Sub AutoCloseMsgBox() 'использую функции API. Для этого объявил функцию MessageBoxTimeOut из библиотеки «User32» [в Module5_MessageBoxTimeOut]
  
  Const lSeconds As Long = 5
    MessageBoxTimeOut 0, "Имя пользователя компьютера не совпадает с указанным!" & _
    vbNewLine & "Программа будет остановлена." & _
    vbNewLine & "Это окно закроется автоматически через 5 секунд.", "Сообщение", _
    vbInformation + vbOKOnly, 0&, lSeconds * 1000
  
  
  Const lSeconds As Long = 3
  MessageBoxTimeOut 0, "Не найдено." & _
     vbNewLine & "Это окно закроется автоматически через 3 секунды" & _
     vbNewLine & "Программа остановится", "Сообщение от Microsoft Corporation: Мы искали, но...", _
  vbInformation + vbOKOnly, 0&, lSeconds * 1000
  
End Sub
'=============================================================================================================================================
Sub AutoCloseMsgBox_2()
    Const lSeconds As Long = 3
    Dim RezPoiska
    'поискали - не нашли. тогда сообщение:
    If RezPoiska Is Nothing Then
        MessageBoxTimeOut 0, "Не найдено." & _
                    vbNewLine & "Открываю файл ""Реестр2"" для уточнения" & _
                    vbNewLine & "Это окно закроется автоматически через 4 секунды", "РЕЗУЛЬТАТ ПОИСКА", _
                 vbInformation + vbOKOnly, 0&, lSeconds * 1000
    Exit Sub
    End If
End Sub



Function IsBookOpen(wbFullName As String) As Boolean  'Функция проверки открыта или закрыта книга
    Dim iFF As Integer, RetVal As Boolean
    iFF = FreeFile
    On Error Resume Next
    Open wbFullName For Random Access Read Write Lock Read Write As #iFF
    RetVal = (Err.Number <> 0)
    Close #iFF
    IsBookOpen = RetVal
End Function

Sub MsgBoxПРИМЕР()
Да = MsgBox("Вам понравился пример?", vbYesNo, "Примерчик))")
If Да = vbYes Then
    MsgBox "Це добре", vbInformation, "Primer"
Else
    MsgBox "Це погано", vbInformation, "Primer"
End If
End Sub
Sub рПРИМЕР()
    ТекстНапоминания = InputBox("введите текст напоминания", "Запрос данных", " ")
  
    If ТекстНапоминания = "" Then
        MsgBox "Текст не введен, закрытие программы", vbCritical, "Текст напоминания пользователя"
        Exit Sub
    End If
End Sub

 If InputBox("Введите пароль Администратора") <> "123" Then MsgBox "Неправильный пароль": Exit Sub
   iProcedure = InputBox(Prompt:="Введите имя процедуры," & _
   vbCrLf & "которую требуется удалить", title:="Удаление подпрограммы")
   If iProcedure$ = "" Then _
   MsgBox "Вы не указали имя ненужной процедуры", 48, "Ошибка": Exit Sub


Sub ПРИМЕР_77() 'InputBox в заданных координатах
    Isk0 = InputBox("Введите шаблон искомого слова ", , Isk0, 12000, 6000)
End Sub
  
Sub пример55()
  ' Reference: Tools - References - Windows Script Host Object Model
  ' File: C:\Windows\System32\wshom.ocx
  Dim WshShell As IWshRuntimeLibrary.WshShell
  Set WshShell = New IWshRuntimeLibrary.WshShell
  WshShell.Popup "Hi!"
End Sub

Sub пример56()
               ' в Офисе 19 сообщение закрывается:
               
Dim WshShell As IWshRuntimeLibrary.WshShell
    Set WshShell = New IWshRuntimeLibrary.WshShell
    WshShell.Popup "Сохранил...", 2, "Сообщение о резервном копировании файла", 48
    
               '  в Офисе 19 сообщение НЕ закрывается:
               
'   CreateObject("WScript.Shell").Popup "Сохранил...", 2, "Сообщение о резервном копировании файла", 48
End Sub


