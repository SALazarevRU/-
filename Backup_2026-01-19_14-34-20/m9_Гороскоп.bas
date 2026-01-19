Attribute VB_Name = "m9_Гороскоп"
'Sub ГороскопСегодня()
Sub ГороскопСегодняТелец(control As IRibbonControl)
   Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev"
'   Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "Хозяин"
       username = Environ("UserName")  ' Получаем имя пользователя.
       If username = SpecifiedUserName Then
            Dim URL, a, b, C, d As String, HTMLDoc, elements, element As Object, i As Integer
            
            Application.ScreenUpdating = False
           
            On Error Resume Next
                If Worksheets(sName) Is Nothing Then
                    Worksheets.Add.Name = "Гороскоп"
                End If
            On Error GoTo 0
            
            On Error GoTo ErrHandler
             
            URL = "https://my-calend.ru/goroskop/telec"
            Set ws = ActiveWorkbook.Sheets("Гороскоп")
            rowNum = 1
            
            Set HTMLDoc = CreateObject("HTMLFile")
            With CreateObject("MSXML2.XMLHTTP")
                .Open "GET", URL, False
                .send
                HTMLDoc.body.innerHTML = .responseText
            End With
            
            Set elements = HTMLDoc.getElementsByTagName("p")
            For Each element In elements
                ws.Cells(rowNum, 1).Value = element.innerText
                rowNum = rowNum + 1
            Next element
            
            a = Range("A1")
            b = Range("A2")
            C = Range("A3")
            d = Range("A4")
            
           Application.DisplayAlerts = False ' Отключить диалоговое окно отображение предупреждений и сообщений
           
           ActiveWorkbook.Sheets("Гороскоп").Delete
           
            Application.DisplayAlerts = True ' Снова включить отображение предупреждений и сообщений
            Application.ScreenUpdating = True
            
            MsgBox d & vbNewLine & vbNewLine & C & vbNewLine & vbNewLine & b & vbNewLine & vbNewLine & a, , "Гороскоп на сегодня       https://my-calend.ru/goroskop/telec"
             
            Set HTMLDoc = Nothing
            Set elements = Nothing
            
       Exit Sub
ErrHandler:
             MsgBox "Неизвестная ошибка", vbExclamation, "Ошибка выполнения"
            On Error GoTo 0
          
    Else
        CreateObject("WScript.Shell").Popup ("В настоящее время доступ заблокирован." & _
                vbNewLine & "Это окно сейчас закроется "), 1, "Информация о блокировке доступа к выполнению программы", 48
      Exit Sub
    End If
End Sub

Sub ГороскопСегодняЗнак(control As IRibbonControl)
' MsgBox "Не определён Знак", vbExclamation, "Ошибка выполнения": Exit Sub
    Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev"
'   Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "Хозяин"
       username = Environ("UserName")  ' Получаем имя пользователя.
       If username = SpecifiedUserName Then
            Dim URL, a, b, C, d As String, HTMLDoc, elements, element As Object, i As Integer
            
            Application.ScreenUpdating = False
           
            On Error Resume Next
                If Worksheets(sName) Is Nothing Then
                    Worksheets.Add.Name = "Гороскоп"
                End If
            On Error GoTo 0
            
            On Error GoTo ErrHandler
             
            URL = "https://my-calend.ru/goroskop/vesy"
            Set ws = ActiveWorkbook.Sheets("Гороскоп")
            rowNum = 1
            
            Set HTMLDoc = CreateObject("HTMLFile")
            With CreateObject("MSXML2.XMLHTTP")
                .Open "GET", URL, False
                .send
                HTMLDoc.body.innerHTML = .responseText
            End With
            
            Set elements = HTMLDoc.getElementsByTagName("p")
            For Each element In elements
                ws.Cells(rowNum, 1).Value = element.innerText
                rowNum = rowNum + 1
            Next element
            
            a = Range("A1")
            b = Range("A2")
            C = Range("A3")
            d = Range("A4")
            
           Application.DisplayAlerts = False ' Отключить диалоговое окно отображение предупреждений и сообщений
           
           ActiveWorkbook.Sheets("Гороскоп").Delete
           
            Application.DisplayAlerts = True ' Снова включить отображение предупреждений и сообщений
            Application.ScreenUpdating = True
            
            MsgBox d & vbNewLine & vbNewLine & C & vbNewLine & vbNewLine & b & vbNewLine & vbNewLine & a, , "Гороскоп на сегодня       https://my-calend.ru/goroskop/telec"
             
            Set HTMLDoc = Nothing
            Set elements = Nothing
            
       Exit Sub
ErrHandler:
             MsgBox "Неизвестная ошибка", vbExclamation, "Ошибка выполнения"
            On Error GoTo 0
          
    Else
        CreateObject("WScript.Shell").Popup ("В настоящее время доступ заблокирован." & _
                vbNewLine & "Это окно сейчас закроется "), 1, "Информация о блокировке доступа к выполнению программы", 48
      Exit Sub
    End If
End Sub

Sub ScrapeDataToMessageBox(Optional Dummy)
    Dim URL As String: URL = "https://my-calend.ru/goroskop/telec"
    Dim HTMLDoc As Object, elements As Object, element As Object
    Dim myTxt
    Dim ares(), lcnt&, lr&
    
    Set HTMLDoc = CreateObject("HTMLFile")
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", URL, False
        .send
        HTMLDoc.body.innerHTML = .responseText
    End With
    
    Set elements = HTMLDoc.getElementsByTagName("p")
    ReDim ares(1 To elements.Length)
    For Each element In elements
        lcnt = lcnt + 1
        ares(lcnt) = element.innerText
'         myTxt = element.innerText
'                Const lSeconds As Long = 4
'                MessageBoxTimeOut 0, myTxt & _
'                vbNewLine, "Гороскоп на сегодня", _
'                vbInformation + vbOKOnly, 0&, lSeconds * 1000
    Next element
    'в обратном порядке
    For lr = UBound(ares) To 1 Step -1
        MsgBox ares(lr)
    Next
Stop 'хххххххххххххххххххххххххххххххххххххххххххххххххххххххххххххххххххххххх
    'исключаем 2-ой по порядку
    For lr = 1 To UBound(ares)
        If lr <> 2 Then
            MsgBox ares(lr)
        End If
    Next
    'собираем в одну строку в обратном порядке
    For lr = UBound(ares) To 1 Step -1
        myTxt = myTxt & vbNewLine & ares(lr)
    Next
    MsgBox myTxt
End Sub
    

