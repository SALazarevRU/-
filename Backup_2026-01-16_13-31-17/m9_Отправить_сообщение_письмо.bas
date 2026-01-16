Attribute VB_Name = "m9_Отправить_сообщение_письмо"
Option Explicit

Sub Отправить_Письмо()
    Dim EmailApp As Outlook.Application
    Dim Source As String
    Set EmailApp = New Outlook.Application
    Dim EmailItem As Outlook.MailItem
    Set EmailItem = EmailApp.CreateItem(olMailItem)
    EmailItem.To = "s.a.lazarev@yandex.ru"
    EmailItem.cc = " "
    EmailItem.BCC = " "
    EmailItem.Subject = "Тема письма"
    EmailItem.HTMLBody = "Добрый день," & "<br>" & "Для Вас сообщение о новом инциденте" & "<br>" & "ФИО," & "<br>" & "Должность"
    Source = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\1234.ini"
    EmailItem.Attachments.Add Source
    EmailItem.send
End Sub

Sub Отправить_сообщение()
    Dim sDateTimeStamp As String
    sDateTimeStamp = VBA.Format(VBA.Now, "  yyyy.mm.dd    HH:MM:SS")
    Dim EmailApp As Outlook.Application
    Dim Source As String
    Set EmailApp = New Outlook.Application
    Dim EmailItem As Outlook.MailItem
    Set EmailItem = EmailApp.CreateItem(olMailItem)
    EmailItem.To = "s.a.lazarev@yandex.ru"
    EmailItem.Subject = "Попытка несанкционированного доступа  " & sDateTimeStamp
    EmailItem.HTMLBody = "Добрый день," & "<br>" & "Для Вас сообщение о новом инциденте" & sDateTimeStamp & "<br>" & "ФИО," & "<br>" & "Должность"
    Source = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\1234.ini"
    EmailItem.Attachments.Add Source
    EmailItem.send
End Sub

Sub ОтправитьФайлПоПочте(control As IRibbonControl) 'Пример макроса для отправки письма с вложением:
    Dim EmailApp As Outlook.Application
    Dim Source, Комментарий As String
    Set EmailApp = New Outlook.Application
    Dim EmailItem As Outlook.MailItem
    
    Комментарий = InputBox("Здесь можно ввести текст комментария к письму", "Запрос комментария")
  
    If Комментарий <> "" Then
      GoTo Skip
    End If
  
Skip:
    Set EmailItem = EmailApp.CreateItem(olMailItem)
    EmailItem.To = "s.lazarev@bsv.legal"
    EmailItem.cc = ""
    EmailItem.BCC = ""
    EmailItem.Subject = "Отправка файла" & "   " & ActiveWorkbook.Name
    EmailItem.HTMLBody = "Добрый день," & "<br>" & "Отправил Вам файл " & ActiveWorkbook.Name & " (во вложении)." & "<br>" & Комментарий & "<br>" & "<br>" & "<br>" & "С уважением: " & "<br>" & "Сергей Лазарев" & "<br>" & "Специалист группы ..."
    Source = ActiveWorkbook.FullName
    EmailItem.Attachments.Add Source
    EmailItem.send
End Sub

Sub Отправить_Надстройку2_по_почте(control As IRibbonControl)
    Dim EmailApp As Outlook.Application
    Dim Source, Комментарий As String
    Set EmailApp = New Outlook.Application
    Dim EmailItem As Outlook.MailItem
    shell ("OUTLOOK")
    
    Комментарий = InputBox("Здесь можно ввести текст комментария к письму", "Запрос комментария")
  
    If Комментарий <> "" Then
      GoTo Skip
    End If
  
Skip:
    Set EmailItem = EmailApp.CreateItem(olMailItem)
    EmailItem.To = "s.lazarev@bsv.legal"
    EmailItem.cc = ""
    EmailItem.BCC = ""
    EmailItem.Subject = "Отправка файла" & "   " & ThisWorkbook.Name
    EmailItem.HTMLBody = "Добрый день," & "<br>" & "Отправил Вам файл " & ThisWorkbook.Name & " (во вложении)." & "<br>" & Комментарий & "<br>" & "<br>" & "<br>" & "С уважением, " & "<br>" & "Лазарев Сергей Александрович" & "<br>" & "Специалист отдела Архива длительного хранения" & "<br>" & "ООО ПКО ""Бюро судебного взыскания""" & "<br>" & "E-mail: s.lazarev@bsv.legal"
    Source = ("C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\Надстройка2.xlam") 'Адрес вложения
    EmailItem.Attachments.Add Source
    EmailItem.send
'    Application.Wait Now + TimeValue("00:00:08")
End Sub
