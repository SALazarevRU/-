Attribute VB_Name = "m9_Открыть_Справку"

Public Sub ОткрытьГайдПоКодингу(control As IRibbonControl)
    Dim helpFilePath As String
    
     If InputBox("Введите пароль Администратора", "Запрос авторизации Пользователя") <> "123" Then
            Да = MsgBox("Пароль не введён или не верен, доступ не предоставлен." & vbNewLine & "Отравить сообщение об инциденте разработчику?", vbYesNo, "Информация о блокировке доступа к выполнению программы")
               If Да = vbYes Then
                 Dim EmailApp As Outlook.Application
                    Dim Source As String
                    Set EmailApp = New Outlook.Application
                    Dim EmailItem As Outlook.MailItem
                    Set EmailItem = EmailApp.CreateItem(olMailItem)
                    EmailItem.To = "s.lazarev@bsv.legal"
                    EmailItem.cc = " "
                    EmailItem.BCC = " "
                    EmailItem.Subject = "Test Email From Excel VBA"
                    EmailItem.HTMLBody = "Добрый день," & "<br>" & "Для Вас сообщение о новом инциденте" & "<br>" & "ФИО," & "<br>" & "Должность"
                '    Source = "" 'ThisWorkbook.FullName
                '    EmailItem.Attachments.Add Source
                    EmailItem.send
                  MsgBox "Сообщение отправлено", vbInformation, "Информация о блокировке доступа к выполнению программы"
                  Exit Sub
                Else
                  MsgBox "Программа будет закрыта", vbInformation, "Информация о состоянии выполнения программы"
                  Exit Sub
               End If
     End If
        helpFilePath = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns" & "\2.chm"
        HtmlHelp 0, helpFilePath, 0, 0
End Sub


Public Sub ВводныйКурсВ_VB(control As IRibbonControl)
    Dim helpFilePath As String
'    2.     Попытка выполнения запроса пользователем, у которого установлены ограничения доступа на уровне записей
     If InputBox("Введите пароль Администратора", "Аторизация") <> "123" Then
    Const lSeconds As Long = 3
                MessageBoxTimeOut 0, "Неверно настроены права в профиле пользователя," & _
                vbNewLine & "в доступе отказано!" & _
                vbNewLine & " " & _
                vbNewLine & "Это окно закроется через 3 секунды.", "Проверка прав доступа", _
                vbInformation + vbOKOnly, 0&, lSeconds * 1000: Exit Sub
   End If
        helpFilePath = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns" & "\1.chm"
        HtmlHelp 0, helpFilePath, 0, 0
End Sub

Public Sub ИнструкцияПоВалидации(control As IRibbonControl)
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Set wApp = CreateObject("Word.Application") 'Создать новый экземпляр приложения Word
    wApp.Visible = True  ' Сделать окно приложения Word видимым
    Set wDoc = wApp.Documents.Open("C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\Справка_ВалидацияПД\Инструкция_Валидация досье должников.docx")
End Sub

