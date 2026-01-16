Attribute VB_Name = "m9_Динамическое_Меню"
Public Логин As String
Public ПарольПО As String

'Создать новый лист
Sub Сообщение0000000000000(control As IRibbonControl)
    Dim sheet As Worksheet
   Dim cell As Range
   Dim sName As String 'Создаём переменную, в которую поместим имя листа.
   sName = "Валидация" 'Помещаем в переменную имя листа
   
   Application.EnableEvents = False 'Отключаем отслеживание событий
   
   On Error Resume Next
   If Worksheets(sName) Is Nothing Then  'действия, если листа нет
       Worksheets.Add.Name = "Валидация"
   End If
   Worksheets("Валидация").UsedRange.ClearContents
   
  Application.EnableEvents = True
End Sub



Sub Сообщение1(control As IRibbonControl)
    Dim URL As String, Wb As Workbook
    Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev"
       username = Environ("UserName")  ' Получаем имя пользователя.
        If username = SpecifiedUserName Then
           shell """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe""" & "https://vremya-ne-zhdet.ru/vba-excel/soderzhaniye-rubriki/"
        Else
           CreateObject("WScript.Shell").Popup ("Нет доступа." & _
                vbNewLine & "Это окно сейчас закроется "), 1, "Информация о блокировке доступа к выполнению программы", 48
            Exit Sub
        End If
End Sub

'Кнопка2 (компонент: button, атрибут: onAction), 2007
Sub ЯндексПоиск(control As IRibbonControl)
'Sub ЯндексПоиск()
    Dim URL As String, Wb As Workbook

    Application.DisplayAlerts = False
    Set Wb = Workbooks.Open("C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\Авторизация.xlsx", False)
    ' Выполнение операций с книгой, например:
    Логин = Worksheets("Авторизация").Range("B2").Value
    ПарольПО = Worksheets("Авторизация").Range("B3").Value
    
    ' Закрытие книги без отображения на экране:
    Wb.Close False  ' (второй параметр — SaveChanges, False означает «не сохранять изменения перед закрытием»)
    Set Wb = Nothing
    
    Application.DisplayAlerts = True
    shell """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe""" & "https://yandex.ru/search/?text=vba+excel....&lr=65&clid=2411726%2F/"" & ActiveCel.Value, vbNormalFocus"
    
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
       .SetText Логин: .PutInClipboard
    End With
'    URL = "https://yandex.ru/search/?text=vba+excel....&lr=65&clid=2411726%2F/"
'    ShellExecute 0, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus
            Start = Timer  ' Определяем время старта
        Do While Timer < Start + 2
            DoEvents  ' Уступаем другим процессам
        Loop

    Копи_Логин_Пароль.Show
    Do While Timer < Start + 1
            DoEvents  ' Уступаем другим процессам
        Loop
 
        Workbooks("Надстройка2.xlam").Activate 'вытягиваем на первый план
End Sub

'Кнопка3 (компонент: button, атрибут: onAction), 2007
Sub Сообщение3(control As IRibbonControl)
    MsgBox "Открыть переводчик в браузере Edge?"
     Dim URL As String, Wb As Workbook
      shell """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe""" & "https://translate.yandex.ru/?from=tabbar&source_lang=en&target_lang=ru/"" & ActiveCel.Value, vbNormalFocus"
    
End Sub
Sub Сообщение4(control As IRibbonControl)
    MsgBox "Был выбран пункт 4"
End Sub
Sub Сообщение5(control As IRibbonControl)
    MsgBox "Был выбран пункт 5"
End Sub
Sub Сообщение6(control As IRibbonControl)
    MsgBox "Был выбран пункт 6"
End Sub

'ДинамическоеМеню1 (компонент: dynamicMenu, атрибут: getContent), 2007
Sub ВернутьДинамическоеМеню(control As IRibbonControl, ByRef content)
Dim sXML As String
    sXML = "<menu itemSize=""normal"" xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" & vbCrLf
    sXML = sXML & "<button id=""Кнопка1"" label=""Программинг"" image=""ЯНДЕКС-2"" description=""Ла ла ла"" onAction = ""Сообщение1""/>" & vbCrLf
    sXML = sXML & "<button id=""Кнопка2"" label=""ЯндексПоиск"" image=""ЯНДЕКС-2"" description=""Ла ла ла"" onAction = ""ЯндексПоиск""/>" & vbCrLf
    sXML = sXML & "<button id=""Кнопка3"" label=""Пункт 3"" description=""Пункт динамического меню"" onAction = ""Сообщение3""/>" & vbCrLf
    
    sXML = sXML & "<button id=""Кнопка4"" label=""Пункт 4"" description=""Пункт динамического меню"" onAction = ""Сообщение4""/>" & vbCrLf
    sXML = sXML & "<button id=""Кнопка5"" label=""Пункт 5"" description=""Пункт динамического меню"" onAction = ""Сообщение5""/>" & vbCrLf
    sXML = sXML & "<button id=""Кнопка6"" label=""Пункт 6"" description=""Пункт динамического меню"" onAction = ""Сообщение6""/>" & vbCrLf
    content = sXML & "</menu>"
End Sub
