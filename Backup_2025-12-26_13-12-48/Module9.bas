Attribute VB_Name = "Module9"

Sub Search_internet77() ' МОЯ
    Dim Добавка_к_URL As String, MyTarget As String: Добавка_к_URL = "sometext"
    MyTarget = Подстановка_Добавки_в_URL(Добавка_к_URL)
End Sub

'При передаче по значению в функцию передаётся копия значения переменной. ByVal
'Любые изменения, сделанные с параметром внутри функции, не затрагивают исходную переменную

Function Подстановка_Добавки_в_URL(ByVal Добавка_к_URL As String)  ' МОЯ
   Dim URL_Start As String, BrowserPath As String
   
    BrowserPath = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"""
    URL_Start = "https://yandex.ru/search/?text=vba+excel...."
    myURL = """" & URL_Start & Добавка_к_URL & """"
    If MsgBox("Открыть[?]: " & myURL, vbYesNo) = vbNo Then Exit Function
    shell BrowserPath & myURL, vbNormalFocus
End Function
