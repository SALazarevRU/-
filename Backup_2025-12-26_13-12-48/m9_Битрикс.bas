Attribute VB_Name = "m9_Битрикс"
'Парсинг содержимого тегов
'Извлечение содержимого тегов с помощью метода getElementsByTagName объекта HTMLFile:
Public myTxt2 As String


Sub ЗапуститьМониторинг(control As IRibbonControl)
'Sub ЗапуститьМониторинг()
If MsgBox("Вы действительно хотите запустить мониторинг активации ПО в рабочие дни?", vbYesNo) = vbNo Then Exit Sub
    ' Запускаем первый цикл мониторинга
    Call ПроверкаЦелевогоВремени_Дек
    MsgBox "Мониторинг времени запущен. Проверка каждые 60 секунд.", vbInformation
End Sub


Sub ПроверкаЦелевогоВремени_Дек()
    Dim сейчас As Date: сейчас = Now                      ' дата + время
    
    ' будний день?
    If Weekday(сейчас, vbMonday) > 5 Then GoTo Планируем
    
    ' нужный час и минута?
    If Hour(сейчас) = 6 And Minute(сейчас) = 36 Then
         Call БитриксСТАРТ
         Call ОстановитьМониторингБТР
        ' больше не планируем, если нужен только один пуск
        Exit Sub
    End If
    
Планируем:
    Application.OnTime сейчас + TimeSerial(0, 1, 0), "ПроверкаЦелевогоВремени_Дек"
End Sub



Sub ОстановитьМониторинг(control As IRibbonControl)
    Call ОстановитьМониторингБТР  ' Передаем управление базовой процедуре
End Sub

Sub ОстановитьМониторингБТР()
    On Error Resume Next ' Игнорируем ошибку, если задача уже отменена
    Application.OnTime _
        EarliestTime:=Time + TimeValue("00:01:00"), _
        Procedure:="ПроверкаЦелевогоВремени_Дек", _
        Schedule:=False
    On Error GoTo 0
'    MsgBox "Мониторинг остановлен.", vbInformation

    ' Запускаем всплывающее окно
'    CreateObject("WScript.Shell").Popup ("Мониторинг остановлен."), 1, "Это окно закроется через 1 секунду", 48
    Application.SendKeys "{NUMLOCK}"  'Активируем правый цифровой блок
End Sub




Public Sub БитриксСТАРТ()
    Dim strFile_Path As String
    Dim username As String
        username = Environ("UserName")
        
    Call БитриксОБНОВИТЬ

     On Error GoTo ErrHandler
     Application.Wait Now + TimeValue("00:00:02")
     
        SetCursorPos 1705, 292  'позиция
        mouse_event &H2, 0, 0, 0, 0 'клик
                      Sleep (300)
                      mouse_event &H4, 0, 0, 0, 0
                      Sleep (300)
        Application.Wait (Now() + TimeValue("00:00:01"))
    'Stop
           
         SetCursorPos 1703, 248 'позиция
        mouse_event &H2, 0, 0, 0, 0 'клик ещё раз
                      Sleep (300)
                      mouse_event &H4, 0, 0, 0, 0
                      Sleep (300)
        Application.Wait (Now() + TimeValue("00:00:03"))
        
    ' Stop
    
        SetCursorPos 1111, 575  'позиция
        mouse_event &H2, 0, 0, 0, 0 'клик
                      Sleep (300)
                      mouse_event &H4, 0, 0, 0, 0
                      Sleep (300)
                      Range("AP1") = "Запущено: " & Now
    'Stop
         Application.SendKeys "{NUMLOCK}"  'Активируем правый цифровой блок
        Exit Sub
    
ErrHandler:
                 strFile_Path = "C:\Users\s.lazarev\Desktop\Лог отчета Бi.txt"
            Open strFile_Path For Append As #1
            Print #1, " "
            Print #1, "ТестерБитриксНачало[]: " & " Ошибка: " & Err.Description & " " & Now   '
            Close #1
     On Error GoTo 0
    Exit Sub
End Sub



Public Sub БитриксОБНОВИТЬ() 'обновление страницы битрикс браузера Яндекс
 Application.Wait Now + TimeValue("00:00:02")
 
   SetCursorPos 110, 15  'позиция вкладки страницы битрикс в браузере Яндекс
   mouse_event &H2, 0, 0, 0, 0 'клик на вкладку страницы битрикс
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
    Application.SendKeys ("^r") 'обновление страницы битрикс в браузере Яндекс
'Stop
    Application.Wait (Now() + TimeValue("00:00:08"))
    SetCursorPos 1170, 550  'позиция на кнопке входа
    mouse_event &H2, 0, 0, 0, 0 'клик по кнопке входа
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
'Stop
    Application.Wait (Now() + TimeValue("00:00:09"))
    SetCursorPos 1870, 100  'позиция моей Аватарки
    mouse_event &H2, 0, 0, 0, 0 'клик по моей Аватарке
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
    
     Application.SendKeys "{NUMLOCK}"  'Активируем правый цифровой блок
End Sub







'++++++++++++++++  ПРОЦЕДУРЫ ОТКЛЮЧЕНИЯ  +++++++++++++++++++++++++++






Sub ЗапуститьМониторингОткл(control As IRibbonControl)
'Sub ЗапуститьМониторинг()
If MsgBox("Вы действительно хотите запустить мониторинг Деактивации ПО в рабочие дни?", vbYesNo) = vbNo Then Exit Sub
    ' Запускаем первый цикл мониторинга
    Call ПроверкаЦелевогоВремени_Откл
    MsgBox "Мониторинг времени Отключения запущен. Проверка каждые 60 секунд.", vbInformation
End Sub


Sub ПроверкаЦелевогоВремени_Откл()
    Dim сейчас As Date: сейчас = Now                      ' дата + время
    
    ' будний день?
    If Weekday(сейчас, vbMonday) > 5 Then GoTo Планируем
    
    ' нужный час и минута?
    If Hour(сейчас) = 16 And Minute(сейчас) = 0 Then
         Call БитриксФИНИШ
         Call ОстановитьМониторингОтклБТР
        ' больше не планируем, если нужен только один пуск
        Exit Sub
    End If
    
Планируем:
    Application.OnTime сейчас + TimeSerial(0, 1, 0), "ПроверкаЦелевогоВремени_Откл"
End Sub


Sub ОстановитьМониторингОткл(control As IRibbonControl)
    Call ОстановитьМониторингОтклБТР  ' Передаем управление базовой процедуре
End Sub


Sub ОстановитьМониторингОтклБТР()
    On Error Resume Next ' Игнорируем ошибку, если задача уже отменена
    Application.OnTime _
        EarliestTime:=Time + TimeValue("00:01:00"), _
        Procedure:="ПроверкаЦелевогоВремени_Откл", _
        Schedule:=False
    On Error GoTo 0
'    MsgBox "Мониторинг остановлен.", vbInformation

    ' Запускаем всплывающее окно
'    CreateObject("WScript.Shell").Popup ("Мониторинг остановлен."), 1, "Это окно закроется через 1 секунду", 48
    Application.SendKeys "{NUMLOCK}"  'Активируем правый цифровой блок
End Sub


Public Sub БитриксОБНОВИТЬдляОткл()
   Application.Wait Now + TimeValue("00:00:02")
 
   SetCursorPos 110, 15  'позиция вкладки страницы битрикс в браузере Яндекс
   mouse_event &H2, 0, 0, 0, 0 'клик на вкладку страницы битрикс
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
    Application.SendKeys ("^r") 'обновление страницы битрикс в браузере Яндекс
'Stop

End Sub



Public Sub БитриксФИНИШ()
    Dim strFile_Path As String
    Dim username As String
        username = Environ("UserName")
        
    Call БитриксОБНОВИТЬдляОткл

     On Error GoTo ErrHandler
     
      Application.Wait (Now() + TimeValue("00:00:05"))
    SetCursorPos 1870, 100  'позиция моей Аватарки
    mouse_event &H2, 0, 0, 0, 0 'клик по моей Аватарке
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
     
'    Stop
     
     Application.Wait Now + TimeValue("00:00:05")
     
        SetCursorPos 1735, 292  'позиция между кнопками "Безопасность" и "Завершить"
        mouse_event &H2, 0, 0, 0, 0 'клик
                      Sleep (300)
                      mouse_event &H4, 0, 0, 0, 0
                      Sleep (300)
        Application.Wait (Now() + TimeValue("00:00:02"))
'    Stop
           
         SetCursorPos 1735, 248 'позиция кнопки "Завершить"
        mouse_event &H2, 0, 0, 0, 0 'клик ещё раз
                      Sleep (300)
                      mouse_event &H4, 0, 0, 0, 0
                      Sleep (300)
        Application.Wait (Now() + TimeValue("00:00:02"))
        
'     Stop
    
        SetCursorPos 1745, 580  'позиция ссылки "Выйти"
        mouse_event &H2, 0, 0, 0, 0 'клик
                      Sleep (300)
                      mouse_event &H4, 0, 0, 0, 0
                      Sleep (300)
                      Range("AP2") = "Отключено: " & Now
    'Stop
    
         Application.SendKeys "{NUMLOCK}"  'Активируем правый цифровой блок
         
     Call PlayWavAPI_Otklychenie

         
        Exit Sub
    
ErrHandler:
                 strFile_Path = "C:\Users\s.lazarev\Desktop\Лог отчета Бi.txt"
            Open strFile_Path For Append As #1
            Print #1, " "
            Print #1, "ТестерБитриксНачало[]: " & " Ошибка: " & Err.Description & " " & Now   '
            Close #1
     On Error GoTo 0
    Exit Sub
End Sub





'++++++++++++++++  ДАЛЬШЕ (ниже) НАДО ПОЧИСТИТЬ ШТОЛЕ КОД  +++++++++++++++++++++++++++





Sub ЗапускBTR(control As IRibbonControl)
'Sub ЗапускBTR()
If MsgBox("Запустить старт по времени: БитриксНачало?", vbYesNo) = vbNo Then Exit Sub

        Application.OnTime TimeValue("06:13:00"), "БитриксНачало"
        
End Sub

Public Sub Запустить_проги_по_времени()

If MsgBox("Запустить старт по времени: БитриксЗавершение, ЗаполнитьОтчетФабрика, ТестерБитриксНачало?", vbYesNo) = vbNo Then Exit Sub

        Application.OnTime TimeValue("17:02:00"), "БитриксЗавершение"
            
        Application.OnTime TimeValue("05:42:00"), "ЗаполнитьОтчетФабрика"
        
        Application.OnTime TimeValue("06:15:00"), "БитриксНачало"
End Sub


Sub БитриксРабочееВремя(control As IRibbonControl)
    Call ParserБитрикс24
    MsgBox myTxt1 & vbNewLine & vbNewLine & myTxt2, , "Битрикс-24: Данные рабочего времени сорудника"
End Sub


Public Sub ВремяРаботыБитрикс24()
 Dim sName As String 'Создаём переменную, в которую поместим имя листа.
 sName = "битр" 'Помещаем в переменную имя листа.
'If MsgBox("Пароль прежний? (Иначе - получите ошибку.)", vbYesNo) = vbNo Then Exit Sub
    'Начните с создания экземпляра объекта. Этот объект будет обрабатывать HTTP-запрос и ответ
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    'Настройка запроса. Далее настройте запрос, указав метод HTTP, URL и все необходимые заголовки.
    'Для базовой аутентификации HTTP необходимо включить заголовок.
    'Учетные данные (имя пользователя и пароль) должны быть закодированы в соответствии с base64.
    Dim username As String
    Dim password As String
    Dim authHeader As String
    
    username = "s.lazarev"
    password = "Qwerty047"
    
    ' Кодируем учетные данные
    Dim credentials As String
    credentials = username & ":" & password
    
    ' Base64 кодирует учетные данные
    Dim base64Credentials As String
    base64Credentials = Base64Encode(credentials)
    
    ' Установите заголовок авторизации
    http.Open "GET", "https://bitrix24.bsv.legal", False
    http.setRequestHeader "Authorization", "Basic " & base64Credentials
    'Отправка запроса. После настройки запроса вы можете отправить его и обработать ответ:
    http.send
    ' Проверяем статус ответа
        If http.status = 200 Then
            Dim pShell As Object
            Set pShell = CreateObject("WScript.Shell")
'            CreateObject("WScript.Shell").Popup "Статус ответа - ОК.", 1, "Вы авторизованы", 64
            Set pShell = Nothing
        Else
            MsgBox "Error: " & http.status & " - " & http.statusText
        End If
        
    Dim trimString As String
    Dim myHtml As String, myFile As Object, myTag As Object, myTxt As String, innerText As String
    
    myHtml = GetHTML1("https://bitrix24.bsv.legal/timeman/timeman.php?login=yes/")
    Set myFile = CreateObject("HTMLFile")
    myFile.body.innerHTML = myHtml
    
   
    
''''''    Set myTag1 = myFile.getElementsByTagName("tr")
    
     Set myTag = myFile.getElementsByTagName("tr")
    myTxt1 = myTag(10).innerText ' .....сменил пароль!?
    myTxt2 = myTag(11).innerText
'     Set myTag = myFile.getElementsByTagName("td")
'     myTxt3 = myTag(75).innerText ' .....сменил пароль!?
'    myTxt4 = myTag(76).innerText
'    myTxt = myTag(86).innerText   '86 = 2025-02-15   87 = 2025-02-18
'    myTxt1 = myTag(82).innerText
'    res = Split(myTxt, "Tags")(0)
'    Debug.Print res
'Debug.Print myTxt1
'    Debug.Print myTxt2
'    Debug.Print myTxt1

    On Error Resume Next
        If Worksheets(sName) Is Nothing Then
                    Worksheets.Add.Name = "битр"
                End If
    On Error GoTo 0
                
Range("A2") = myTxt1
Range("A3") = myTxt2
           


MsgBox myTxt1 & vbNewLine & vbNewLine & myTxt2, , "Битрикс-24: Данные рабочего времени сорудника"
 
 
 If Worksheets(sName) Is Nothing Then
                    Worksheets.Add.Name = "битр"
                End If
            myTxt1 = Range("A2")
            myTxt2 = Range("A3")
            
   Dim strFile_Path As String
'   Dim username As String
       username = Environ("UserName")
       strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
   Open strFile_Path For Append As #1
   Print #1, myTxt2 & vbNewLine & Now   '
   Close #1

End Sub



Public Sub ParserБитрикс24()
'If MsgBox("Пароль прежний? (Иначе - получите ошибку.)", vbYesNo) = vbNo Then Exit Sub
    'Начните с создания экземпляра объекта. Этот объект будет обрабатывать HTTP-запрос и ответ
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    'Настройка запроса. Далее настройте запрос, указав метод HTTP, URL и все необходимые заголовки.
    'Для базовой аутентификации HTTP необходимо включить заголовок.
    'Учетные данные (имя пользователя и пароль) должны быть закодированы в соответствии с base64.
    Dim username As String
    Dim password As String
    Dim authHeader As String
    
    username = "s.lazarev"
    password = "Qwerty047"
    
    ' Кодируем учетные данные
    Dim credentials As String
    credentials = username & ":" & password
    
    ' Base64 кодирует учетные данные
    Dim base64Credentials As String
    base64Credentials = Base64Encode(credentials)
    
    ' Установите заголовок авторизации
    http.Open "GET", "https://bitrix24.bsv.legal", False
    http.setRequestHeader "Authorization", "Basic " & base64Credentials
    'Отправка запроса. После настройки запроса вы можете отправить его и обработать ответ:
    http.send
    ' Проверяем статус ответа
        If http.status = 200 Then
            Dim pShell As Object
            Set pShell = CreateObject("WScript.Shell")
'            CreateObject("WScript.Shell").Popup "Статус ответа - ОК.", 1, "Вы авторизованы", 64
            Set pShell = Nothing
        Else
            MsgBox "Error: " & http.status & " - " & http.statusText
        End If
        
     Dim trimString As String
    Dim myHtml As String, myFile As Object, myTag As Object, myTxt As String, innerText As String
    
    myHtml = GetHTML1("https://bitrix24.bsv.legal/timeman/timeman.php?login=yes/")
    Set myFile = CreateObject("HTMLFile")
    myFile.body.innerHTML = myHtml
    
'    Set myTag1 = myFile.getElementsByTagName("tr") 'Так было в рабочем варианте
    
    Set myTag = myFile.getElementsByTagName("td")
    myTxt1 = myTag(86).innerText ' .....сменил пароль!?
    myTxt2 = myTag(11).innerText  ' myTxt = myTag(86).innerText   86 = 2025-02-15   87 = 2025-02-18
'    myTxt1 = myTag(82).innerText
'    res = Split(myTxt, "Tags")(0)
'    Debug.Print res
'Debug.Print myTxt1
'    Debug.Print myTxt2
'    Debug.Print myTxt1
MsgBox myTxt1 & vbNewLine & vbNewLine & myTxt2, , "Битрикс-24: Данные рабочего времени сорудника"

   Dim strFile_Path As String
'   Dim username As String
       username = Environ("UserName")
       strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
   Open strFile_Path For Append As #1
   Print #1, myTxt2 & vbNewLine & Now   '
   Close #1

End Sub

'Парсинг HTML - страниц(MSXML2.xmlHttp)
'Пользовательская функция GetHTML1 (VBA Excel) для извлечения (парсинга) текстового содержимого
'из html-страницы сайта по ее URL-адресу с помощью объекта «msxml2.xmlhttp»:
Function GetHTML1(ByVal myURL As String) As String
On Error Resume Next
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", myURL, False
        .send
        Do: DoEvents: Loop Until .readyState = 4
        GetHTML1 = .responseText
    End With
End Function

'Функция кодирования Base64.Вам понадобится функция для выполнения кодирования Base64. Вот простая реализация:
 Public Function Base64Encode(inData As String) As String
    Dim arrData() As Byte
    arrData = StrConv(inData, vbFromUnicode)
    Dim objXML As Object
    Dim objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    Base64Encode = objNode.text
End Function
'Заключение.
'Использовать базовую проверку подлинности HTTP в Excel VBA очень просто.
'Выполнив описанные выше действия, вы можете безопасно отправлять запросы к API, требующим проверки подлинности.


Sub Битрикс78()
    Dim fff
    fff = FreeFile
    Open "C:\Users\Хозяин\Desktop\HTMLcode.txt" For Input As #fff
  
    pLeft = InStr(1, fff, "<a href=""/company/personal/user/453/""")
    pRight = InStr(pLeft, fff, "data-id=""453""")
    v = Val(Mid(fff, pLeft, pRight - pLeft)) ' rate
    
    Debug.Print v
    
    If IsNumeric(v) Then
        ActiveCell = "+" & v & "°C"
    End If
    Close #fff ' Закрываем файл
     
            Dim MyFSO As New FileSystemObject ' - Смотрим полный ответ из 3500 строк.
            Call shell("C:\Windows\System32\Notepad.exe" & " " & "D:\OneDrive\Браузер\Ответ.txt", vbNormalFocus)
    
     
End Sub
Sub Битрикс80()
   Const ForReading = 1
    Dim fso As Object
    Dim fl As Object
    Dim str As String
    
    

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fl = fso.OpenTextFile("C:\Users\Хозяин\Desktop\HTMLcode.txt", ForReading, False)
    

 Debug.Print v
     
End Sub
Function vvv$(t$)
  vvv = "user/453" & Split(t, "user/453")(1)
End Function




Public Sub БитриксЗавершение()
    Dim strFile_Path As String
    Dim username As String
        username = Environ("UserName")

On Error Resume Next
'''    shell ("C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe -url https://bitrix24.bsv.legal/timeman/timeman.php") ' Авторизация требуется
    shell ("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe -url https://bitrix24.bsv.legal/timeman/timeman.php"), vbNormalFocus
 
 On Error GoTo 0
 On Error GoTo ErrHandler
 
  Application.Wait Now + TimeValue("00:00:08") ' 8 сек ждем пока откроется вкладка и прорисуются элементы
 
  SetCursorPos 1315, 100 'позиция
Stop
               mouse_event &H2, 0, 0, 0, 0 'клик
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
              
    
    Application.Wait Now + TimeValue("00:00:02")
    SetCursorPos 1450, 230 'позиция
'Stop
               mouse_event &H2, 0, 0, 0, 0 'клик
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
               Sleep (300)
     
    Application.Wait Now + TimeValue("00:00:02")
    SetCursorPos 1760, 18 'позиция
    
    Application.SendKeys "{NUMLOCK}"  'Активируем правый цифровой блок
    
    Call ParserБитрикс24
    
        strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
    Open strFile_Path For Append As #1
    Print #1, " "
    Print #1, "Отработал Саб БитриксЗавершение() " & username & " " & Now   '
    Print #1, myTxt2
    Close #1
    
    Exit Sub
    
ErrHandler:
                 strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
            Open strFile_Path For Append As #1
            Print #1, " "
            Print #1, "БитриксЗавершение[]: " & " Ошибка: " & Err.Description & " " & Now   '
            Close #1
     On Error GoTo 0
    Exit Sub

End Sub

Sub Стартер()
If MsgBox("Применить процедуру [ТестерБитриксНачало]?", vbYesNo) = vbNo Then Exit Sub
Application.OnTime TimeValue("13:47:00"), "БитриксНачало"
End Sub

Public Sub БитриксНач() 'КЛИК на кнопку обновления страницы браузеры Эдж и Яндекс
 Application.Wait Now + TimeValue("00:00:02")
 
   SetCursorPos 1460, 18  'позиция
   mouse_event &H2, 0, 0, 0, 0 'клик на панель вкладок
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
    Application.Wait (Now() + TimeValue("00:00:02"))
    SetCursorPos 70, 50  'позиция на кнопке обновления страницы
    mouse_event &H2, 0, 0, 0, 0 'клик
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
    Application.Wait (Now() + TimeValue("00:00:02"))
    SetCursorPos 1460, 18  'позиция
End Sub

Public Sub БитриксНачало()
    Dim strFile_Path As String
    Dim username As String
        username = Environ("UserName")

 On Error GoTo ErrHandler
 Application.Wait Now + TimeValue("00:00:02")
 
    SetCursorPos 1560, 20  'позиция
    mouse_event &H2, 0, 0, 0, 0 'клик
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
    Application.Wait (Now() + TimeValue("00:00:01"))
    
    Application.SendKeys ("^r")
    Application.Wait (Now() + TimeValue("00:00:05"))
    
     SetCursorPos 1865, 105 'позиция
    mouse_event &H2, 0, 0, 0, 0 'клик ещё раз
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
    Application.Wait (Now() + TimeValue("00:00:03"))
    
    
'    Call SendKeys("This is Some Text", True)
'    Application.SendKeys ("^r")
'  Stop
    SetCursorPos 1760, 175  'позиция
    mouse_event &H2, 0, 0, 0, 0 'клик
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
     
    Application.Wait (Now() + TimeValue("00:00:02"))
    mouse_event &H2, 0, 0, 0, 0 'клик на старт рабочего дня
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
'   Stop
    Application.Wait (Now() + TimeValue("00:00:02"))
    SetCursorPos 1550, 220  'позиция
    mouse_event &H2, 0, 0, 0, 0 'клик на начать раб день
                  Sleep (300)
                  mouse_event &H4, 0, 0, 0, 0
                  Sleep (300)
'   Stop
'    Application.SendKeys ("^r")
    Application.Wait (Now() + TimeValue("00:00:01"))
    SetCursorPos 1460, 18  'позиция
    
       
    Application.SendKeys "{NUMLOCK}"  'Активируем правый цифровой блок
    
'    Call ParserБитрикс24

'         strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
'    Open strFile_Path For Append As #1
'    Print #1, " "
'    Print #1, "Отработал Саб ТестерБитриксНачало() " & username & " " & Now   '
'    Print #1, myTxt2
'    Close #1
    
    Exit Sub
    
ErrHandler:
                 strFile_Path = "D:\OneDrive\Лог отчета фабрика.txt"
            Open strFile_Path For Append As #1
            Print #1, " "
            Print #1, "ТестерБитриксНачало[]: " & " Ошибка: " & Err.Description & " " & Now   '
            Close #1
     On Error GoTo 0
    Exit Sub
End Sub

