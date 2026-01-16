Attribute VB_Name = "m9_ѕроверка_открыт_ли_файл"
Option Explicit

Sub checkWB() ' –јЅќ„јя процедура дл€ проверки, открыт ли файл в Excel, если да - на передний план, если нет - открыть и на передний план.
 
Dim Wb As Workbook
Dim myWB As String
Dim FileName As String

FileName = "C:\Users\’оз€ин\Desktop\ќтчет по клаймам за июнь 2025.xlsx"
myWB = "ќтчет по клаймам за июнь 2025.xlsx"
 
For Each Wb In Workbooks
    If Wb.Name = myWB Then
        Wb.Activate
'        MsgBox "Workbook Is Open!"
        Exit Sub
    End If
Next Wb
 Workbooks.Open FileName
'MsgBox "Workbook is not open"
 
End Sub
