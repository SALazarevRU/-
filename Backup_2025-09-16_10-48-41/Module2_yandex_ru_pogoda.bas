Attribute VB_Name = "Module2_yandex_ru_pogoda"
Sub yandex_ru_pogoda()
'Sub ГороскопСегодня(control As IRibbonControl)

    Dim URL, a, b, C, d As String, HTMLDoc, elements, element As Object, i As Integer
    
    Application.ScreenUpdating = False
   
    On Error Resume Next
        If Worksheets(sName) Is Nothing Then
            Worksheets.Add.Name = "Гороскоп"
        End If
    On Error GoTo 0
    
    On Error GoTo ErrHandler
     
    URL = "https://yandex.ru/pogoda/en/novosibirsk?ysclid=mbtv1ybvf9777357906&lat=55.030199&lon=82.92043"
    Set ws = Workbooks("Лист Microsoft Excel (3).xlsx").Sheets("Гороскоп")
    rowNum = 1
    
    Set HTMLDoc = CreateObject("HTMLFile")
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", URL, False
        .send
        HTMLDoc.body.innerHTML = .responseText
    End With
    
    Set elements = HTMLDoc.getElementsByTagName("a")
    For Each element In elements
        ws.Cells(rowNum, 1).Value = element.innerText
        rowNum = rowNum + 1
    Next element
    Exit Sub
    a = Range("A1")
    b = Range("A2")
    C = Range("A3")
    d = Range("A4")
   
    Sheets("Гороскоп").Delete
    Application.ScreenUpdating = True
    
    MsgBox d & vbNewLine & vbNewLine & C & vbNewLine & vbNewLine & b & vbNewLine & vbNewLine & a, , "Гороскоп на сегодня       https://my-calend.ru/goroskop/telec"
     
    Set HTMLDoc = Nothing
    Set elements = Nothing
   
    Exit Sub
ErrHandler:
     MsgBox "Неизвестная ошибка", vbExclamation, "Ошибка выполнения"
    On Error GoTo 0
    Exit Sub
End Sub
