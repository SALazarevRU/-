Attribute VB_Name = "m9_Подсчет_файлов_в_папке"
Option Explicit


'Макрос для подсчета файлов в папке

Private Sub CountFilesInFolder(strDir As String, Optional strType As String)
Dim file As Variant, i As Integer
If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
file = Dir(strDir & strType)
While (file <> "")
i = i + 1
file = Dir
Wend
MsgBox i
End Sub

'Как использовать макрос CountFilesInFolder   C:\Users\s.lazarev\AppData\Local\Temp

'Подсчет txt файлов в папке
Sub Demo()
Call CountFilesInFolder("C:\Users\Ryan\Documents\", "*txt")
End Sub
'Подсчет файлов Excel в папке
Sub Demo2()
Call CountFilesInFolder("C:\Users\Ryan\Documents\", "*.xls*")
End Sub
'Посчитать все файлы в папке
Sub Demo3()
Call CountFilesInFolder("C:\Users\Ryan\Documents\")
End Sub
'Подсчитывайте файлы только с именем файла «report»
Sub Demo4()
Call CountFilesInFolder("C:\Users\Ryan\Documents\", "*report*")
End Sub

Sub СчитатьВпапкеТЕМП()
Call CountFilesInFolder("C:\Users\s.lazarev\AppData\Local\Temp\", "*JPG")
End Sub

'===============================================================================

Sub testDirFunction() ' РАБОЧИЙ МАКРОС для подсчета файлов в папке
    Dim counter
    Dim fn
     
    ChDir ("C:\Users\s.lazarev\AppData\Local\Temp\")
         
    fn = Dir("*.JPG")
    counter = 0
     
    While Len(fn) > 0
     
    counter = counter + 1
    fn = Dir()
    Wend
    If MsgBox("Общее количество файлов .JPG в папке TEMP :    " & counter & "                 " & vbNewLine & "Удалить их?", vbYesNo) = vbNo Then Exit Sub
'    MsgBox "Общее количество файлов .JPG: " & counter & vbNewLine & "Удалить их?", 64, "Total count"
End Sub

