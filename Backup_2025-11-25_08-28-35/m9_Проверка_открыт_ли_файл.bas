Attribute VB_Name = "m9_ѕроверка_открыт_ли_файл"
Option Explicit

Sub checkWB() ' –јЅќ„јя процедура дл€ проверки, открыт ли файл в Excel, если да - на передний план, если нет - открыть и на передний план.
 
Dim wb As Workbook
Dim myWB As String
Dim FileName As String

FileName = "C:\Users\’оз€ин\Desktop\ќтчет по клаймам за июнь 2025.xlsx"
myWB = "ќтчет по клаймам за июнь 2025.xlsx"
 
For Each wb In Workbooks
    If wb.Name = myWB Then
        wb.Activate
'        MsgBox "Workbook Is Open!"
        Exit Sub
    End If
Next wb
 Workbooks.Open FileName
'MsgBox "Workbook is not open"
 
End Sub
