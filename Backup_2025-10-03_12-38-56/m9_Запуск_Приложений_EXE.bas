Attribute VB_Name = "m9_Запуск_Приложений_EXE"
Option Explicit
Public NameEXE2, Procedure As String
'Public Procedure As String

Sub AdobeFireworks(control As IRibbonControl)
    NameEXE2 = "AdobeFireworks"
    If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBoxForStartEXE"
    End If
End Sub

Sub Калькулятор(control As IRibbonControl)
    NameEXE2 = "Калькулятор"
    If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBoxForStartEXE"
    End If
End Sub

Sub Acrobat(control As IRibbonControl)
    NameEXE2 = "Acrobat"
    If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBoxForStartEXE"
    End If
End Sub

Public Sub Оутлук(control As IRibbonControl)
'Public Sub Outlook()
    NameEXE2 = "Outlook"
    If Not gRibbon Is Nothing Then '   Вывод результата в editBox:
           gRibbon.InvalidateControl "editBoxForStartEXE"
    End If
End Sub
Sub Сканер(control As IRibbonControl)
    NameEXE2 = "Сканер"
    If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBoxForStartEXE"
    End If
End Sub

Sub RidNacs(control As IRibbonControl)
    NameEXE2 = "RidNacs"
    If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBoxForStartEXE"
    End If
End Sub

Sub TotalCommander(control As IRibbonControl)
    NameEXE2 = "TotalCommander"
    If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBoxForStartEXE"
    End If
End Sub

Public Sub ИмяПрограммыEXE(editBox As IRibbonControl, ByRef Text)
    Text = " " & NameEXE2
End Sub

'Callback
Sub START_EXE(control As IRibbonControl)
On Error Resume Next
Select Case NameEXE2
Case "AdobeFireworks"
    shell "C:\Program Files (x86)\Adobe\Adobe Fireworks CS6\Fireworks.exe", vbNormalFocus
Case "Калькулятор"
    shell "C:\Windows\System32\calc.exe", vbNormalFocus
Case "Acrobat"
    shell "C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe", vbNormalFocus
Case "Outlook"
    shell ("OUTLOOK")
Case "Сканер"
    shell "C:\Program Files (x86)\Brother\iPrint&Scan\Brother iPrint&Scan.exe", vbNormalFocus
Case "RidNacs"
    shell "C:\Users\s.lazarev\Documents\DISTRIBUTIVE\RidNacs.exe", vbNormalFocus
Case "TotalCommander"
    shell "C:\Users\s.lazarev\Documents\DISTRIBUTIVE\Total Commander 9.50 x64\App\TotalCommander\TOTALCMD64.EXE", vbNormalFocus
Case Else
    MsgBox "Я, дятел, не определил значение переменной  NameEXE2!", vbCritical, "Внимание!"
End Select
End Sub

Function Подстановка_Добавки_в_URL(ByVal Добавка_к_URL As String)
   Dim URL_Start As String, BrowserPath As String
   
    BrowserPath = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"""
    URL_Start = "https://yandex.ru/search/?text=vba+excel...."
    myURL = """" & URL_Start & Добавка_к_URL & """"
    If MsgBox("Открыть[?]: " & myURL, vbYesNo) = vbNo Then Exit Function
    shell BrowserPath & myURL, vbNormalFocus
End Function
