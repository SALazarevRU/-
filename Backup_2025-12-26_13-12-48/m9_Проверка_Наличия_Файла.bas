Attribute VB_Name = "m9_Проверка_Наличия_Файла"
Option Explicit

Sub Проверка_Наличия_Файла()
    Dim strFileName As String
    Dim strFileExists As String
'    strFileName = "C:\Users\Хозяин\Desktop\Документ Microsoft Word99.docx"
    strFileName = "C:\Users\s.lazarev\Desktop\Документ Microsoft Word99.docx"
    strFileExists = Dir(strFileName)
    If strFileExists = "" Then
        MsgBox "Выбранный файл не существует"
    Else
        MsgBox "Выбранный файл существует"
    End If
End Sub
