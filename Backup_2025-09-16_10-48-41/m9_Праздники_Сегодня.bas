Attribute VB_Name = "m9_Праздники_Сегодня"

Sub праздникиСегодня(control As IRibbonControl)
    Dim myHtml As String, myFile As Object, myTag As Object, myTxt As String
On Error GoTo ErrHandler
    myHtml = GetHTML1("https://calend-online.com/holiday/")
    Set myFile = CreateObject("HTMLFile")
    myFile.body.innerHTML = myHtml
    Set myTag = myFile.getElementsByTagName("ul")
    myTxt = myTag(2).innerText
    MsgBox "Праздники сегодня, " & Date & _
     vbNewLine & _
     vbNewLine & myTxt, vbOKOnly, "calend-online.com"
       Exit Sub
ErrHandler:
     MsgBox "Возникла ошибка: " & _
     vbNewLine & Err.Description & " https://calend-online.com/holiday/ " & _
     vbNewLine & _
     vbNewLine & "Программа остановлена", vbExclamation, "Ошибка выполнения"
     On Error GoTo 0
   Exit Sub
End Sub

Function GetHTML1(ByVal myURL As String) As String
On Error Resume Next
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", myURL, False
        .send
        Do: DoEvents: Loop Until .readyState = 4
        GetHTML1 = .responseText
    End With
End Function

Sub Goroskop_1(Optional Dummy)
    Dim myFile As Object, myTag As Object, myTxt As String, myHtml As String, Procedure As String: Procedure = "Goroskop_1"
On Error GoTo ErrHandler

    myHtml = GetHTML1("https://my-calend.ru/goroskop/telec/")
    Set myFile = CreateObject("HTMLFile")
    myFile.body.innerHTML = myHtml
'Debug.Print myHtml
    Set myTag = myFile.getElementsByTagName("div")

    myTxt = myTag(1).innerText
    MsgBox myTxt
    Exit Sub
     
ErrHandler:
     MsgBox "Ошибка: " & Err.Description & _
     vbNewLine & _
     vbNewLine & "Module: " & Application.VBE.ActiveCodePane.CodeModule.Name & _
     vbNewLine & "Procedure: " & Procedure, vbExclamation, "Ошибка выполнения"
    
With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"): .SetText Err.Description: .PutInClipboard
End With  ' На всяк случай отправил в буфер..

    Dim Добавка_к_URL As String, MyTarget As String: Добавка_к_URL = Err.Description
    MyTarget = Подстановка_Добавки_в_URL(Добавка_к_URL)
    On Error GoTo 0
    Exit Sub
End Sub


Function Подстановка_Добавки_в_URL(ByVal Добавка_к_URL As String)
   Dim URL_Start As String, BrowserPath As String
   
    BrowserPath = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"""
    URL_Start = "https://yandex.ru/search/?text=vba+excel...."
    myURL = """" & URL_Start & Добавка_к_URL & """"
    If MsgBox("Открыть[?]: " & myURL, vbYesNo) = vbNo Then Exit Function
    shell BrowserPath & myURL, vbNormalFocus
End Function


Sub праздникиСегодня_2()
Dim myHtml As String, myFile As Object, myTag As Object, myTxt As String
    myHtml = GetHTML1("https://www.rbc.ru/life/prazdniki/kakoj-segodnya-prazdnik")
    Set myFile = CreateObject("HTMLFile")
    myFile.body.innerHTML = myHtml
    Set myTag = myFile.getElementsByTagName("ul")
    myTxt = myTag(2).innerText
    MsgBox myTxt
End Sub

'Общая цель кода
'Таким образом, код предназначен для удаления лишних двойных переносов строк из строки RemoveNumbers2.
'Он будет продолжать заменять все двойные переводы строк на одинарные до тех пор, пока в строке останутся
'двойные переводы строк. В результате этой обработки вы получите строку, в которой все лишние пустые
'строки будут удалены, оставив только одинарные переходы на новую строку между элементами.

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

Sub праздникиСегодня_4()
Dim myHtml As String, myFile As Object, myTag As Object, myTxt As String
    myHtml = GetHTML1("https://prazdniki-segodnya.ru/")
    Set myFile = CreateObject("HTMLFile")
    myFile.body.innerHTML = myHtml
    Set myTag = myFile.getElementsByTagName("div")
    myTxt = myTag(23).innerText
    MsgBox myTxt
End Sub

Private Sub праздникиСегодня_5()
    Dim myHtml As String
    Dim myFile As Object
    Dim myTag As Object
    Dim myTxt As String
    Dim formattedTxt As String
    myHtml = GetHTML1("https://my-calend.ru/holidays?ysclid=mbob1sd69r36482705")
    Set myFile = CreateObject("HTMLFile")
    myFile.body.innerHTML = myHtml
    Set myTag = myFile.getElementsByTagName("ul")
    myTxt = myTag(7).innerText
    formattedTxt = RemoveNumbers2(myTxt)
    MsgBox formattedTxt
End Sub
