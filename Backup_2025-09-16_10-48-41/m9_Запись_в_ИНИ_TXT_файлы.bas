Attribute VB_Name = "m9_Запись_в_ИНИ_TXT_файлы"
Option Explicit

Sub Test777()
Dim ff As Integer
'Получаем свободный номер для открываемого файла
ff = FreeFile
'Открываем (или создаем) файл для чтения и записи
Open ThisWorkbook.Path & "\Log.ini" For Output As ff
'Записываем в файл одну строку
Write #ff, Time
'Закрываем файл
Close ff
'Открываем файл для просмотра
'ThisWorkbook.FollowHyperlink (ThisWorkbook.Path & "\1234.ini")
End Sub

Sub Test778()
Dim ff As Integer, str1 As String, str22 As String
'Получаем свободный номер для открываемого файла
ff = FreeFile
'Открываем файл myFile1.txt для чтения
Open ThisWorkbook.Path & "\1234.ini" For Input As ff
'Считываем строку из файла и записываем в переменные
Input #ff, str1
Close ff
str22 = Replace(str1, "#", "")
'Смотрим, что записалось в переменные
MsgBox "str1 = " & str22
End Sub

Sub Запись2строкВИниФайл()
    Dim ff As Integer
    Dim myТекст As String
    Dim dNow
    dNow = Now
    myТекст = "ДаБлять"
               ff = FreeFile
                Open ThisWorkbook.Path & "\Log.ini" For Output As ff
                Write #ff, dNow
                Write #ff, myТекст
                Close ff
    End Sub
    
Sub ЗаписатьВЛогOPEN()
   Open ThisWorkbook.Path & "\Log.txt" For Append As #1
    Print #1, Date$ + " " + Time$ + "  User: " + Application.username + " " + "открытие файла " + ActiveWorkbook.FullName
    Print #1, ТВоздуха
    Close #1
End Sub
Sub ЗаписатьВЛогCLOSE()
    Open ThisWorkbook.Path & "\Log.txt" For Append As #1
    Print #1, Date$ + " " + Time$ + "  User: " + Application.username + " " + "Закрытие файла " + ActiveWorkbook.FullName
    Print #1, ТВоздуха
    Close #1
End Sub

    
Sub ЗаписьЧтениеЧисткаЛогФайла()
Dim app As Application
Dim i As Integer
Dim User As String
Set app = Excel.Application
    Open ThisWorkbook.Path & "\Log.txt" For Append As #1
    Print #1, app.username + " " + Date$ + " " + Time$
    Print #1, "хххх" + " " + "ДаБлять" + " " + Time$
    Print #1, ТВоздуха
    Close #1
    
   Dim FSO As FileSystemObject
    Dim TextStream As TextStream
    Dim S As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TextStream = FSO.OpenTextFile(ThisWorkbook.Path & "\Log.txt", ForReading)
    S = TextStream.ReadAll
    
    MsgBox (S)
    Set FSO = Nothing
    If MsgBox("Очистить Log.txt?", vbYesNo) = vbNo Then Exit Sub
     Open ThisWorkbook.Path & "\Log.txt" For Output As #1: Close #1 'DeleteFileContent
End Sub


