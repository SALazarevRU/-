Attribute VB_Name = "Module8"
Sub Search_internet()
    Dim Target As String, MyTarget As String
    Target = "sometext"
    MyTarget = Подстановка(ЧтоИщем)
End Sub



Public Function Подстановка(ByVal myURL As String)
   Dim URL_Start As String, BrowserPath As String, MyTarget As String
            BrowserPath = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"""
            URL_Start = "https://yandex.ru/search/?text=vba+excel...."
              myURL = """" & URL_Start & Target & """"
                If MsgBox("Открыть[?]: " & myURL, vbYesNo) = vbNo Then Exit Function
                  shell BrowserPath & myURL, vbNormalFocus
End Function


Function RemoveNumbers2(CellText As String)
 ' Удаление всех цифр
    With CreateObject("VBScript.RegExp")
        .Global = True 'Значение свойства .Global определяет, как ведётся поиск совпадений в строке:
'True — поиск всех возможных совпадений;
        .Pattern = "\d"  'Регулярное выражение для поиска цифр
        RemoveNumbers2 = .Replace(CellText, vbCrLf) ' Удаление всех цифр

    End With
        Do While InStr(1, RemoveNumbers2, vbCrLf & vbCrLf, 1) <> 0
        RemoveNumbers2 = Replace(RemoveNumbers2, vbCrLf & vbCrLf, vbCrLf, vbTextCompare)
        Loop
End Function

Private Sub праздникиСегодня_3()
    Dim myHtml As String
    Dim myFile As Object
    Dim myTag As Object
    Dim myTxt As String
    Dim formattedTxt As String

    ' Получение HTML-кода с веб-страницы
    myHtml = GetHTML1("https://my-calend.ru/holidays?ysclid=mbob1sd69r36482705")
    
    ' Создание HTML-документа
    Set myFile = CreateObject("HTMLFile")
    myFile.body.innerHTML = myHtml
    
    ' Получение списка праздников
    Set myTag = myFile.getElementsByTagName("ul")
    myTxt = myTag(7).innerText
    
    ' Обработка текста: удаление чисел и форматирование
    formattedTxt = RemoveNumbers2(myTxt)
    
    ' Вывод отформатированного текста
    MsgBox formattedTxt
End Sub
