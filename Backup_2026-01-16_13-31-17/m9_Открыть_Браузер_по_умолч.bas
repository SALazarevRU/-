Attribute VB_Name = "m9_Открыть_Браузер_по_умолч"
Sub Открыть_Браузер(Optional Dummy)
ThisWorkbook.FollowHyperlink "https://www.planetaexcel.ru/"
    Dim ИмяПроцедуры As String
    Dim ИмяМодуля As String
    ИмяМодуля = ActiveWorkbook.VBProject.VBE.ActiveCodePane.CodeModule.Parent.Name
    ИмяПроцедуры = "Открыть_Браузер(control As IRibbonControl)"
    On Error GoTo ErrHandler
    
   CreateObject("WScript.Shell").Run "https://ExcelVBA.ru/"
   
    Exit Sub
ErrHandler:     MsgBox "ТЕКСТ"
     MsgBox "Ошибка: " & Err.Description & "https://: " & _
     vbNewLine & "Имя Модуля: " & ИмяМодуля & _
     vbNewLine & "Имя Процедуры: " & ИмяПроцедуры, vbExclamation, "Ошибка выполнения"
    
    On Error GoTo 0  ' Сброс обработчика ошибок
    Exit Sub
End Sub

Sub Открыть_Браузер_2(Optional Dummy)
   ThisWorkbook.FollowHyperlink "https://www.planetaexcel.ru/"
   Application.Wait Now + TimeValue("00:00:01")
End Sub



Sub Открыть_Браузер_3(Optional Dummy)
'ActiveWorkbook.FollowHyperlink Address:="https://www.automateexcel.com/excel", NewWindow:=True
ActiveWorkbook.FollowHyperlink Address:="https://www.automateexcel.com/excel"
End Sub


Sub Открыть_Браузер_4()
    Dim chrome As Object
    Set chrome = CreateObject("Chrome.Application")
    chrome.Navigate "https://bitrix24.bsv.legal/timeman/timeman.php?login=yes"
        With chrome
        .Visible = True
        .Navigate2 "https://bitrix24.bsv.legal/timeman/timeman.php?login=yes"
        .Top = 0
        .Left = 0
        .Height = 600
        .Width = 900
        .resizable = True
    End With
       Application.Wait Now + TimeValue("00:00:04")
End Sub


Sub пример55()
  ' Reference: Tools - References - Windows Script Host Object Model
  ' File: C:\Windows\System32\wshom.ocx
  Dim WshShell As IWshRuntimeLibrary.WshShell
  Set WshShell = New IWshRuntimeLibrary.WshShell
  WshShell.Popup "Hi!"
End Sub



