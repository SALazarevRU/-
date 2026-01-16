Attribute VB_Name = "m9_Запустить_Outlook_Свёрнуто"
Option Explicit
 
Sub ЗапуститьOutlookСвёрнуто()
 
    ' Путь к Outlook - можно уточнить с помощью процедуры что выше "FindOutlookPath", если установлен в другом месте
    Dim outlookPath As String
    outlookPath = """C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"""
 
    ' Создаем объект Shell
    Dim shell       As Object
    Set shell = CreateObject("WScript.Shell")
 
    ' Запускаем Outlook в свернутом виде: параметр 7 — свернутое окно
    shell.Run outlookPath, 7, False
End Sub

Sub CloseOutlook()
    Dim objAppOL As Outlook.Application
    Set objAppOL = GetObject(Class:="Outlook.application")
    objAppOL.Quit
    Set objAppOL = Nothing
End Sub
