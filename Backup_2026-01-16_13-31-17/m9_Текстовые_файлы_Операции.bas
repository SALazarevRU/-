Attribute VB_Name = "m9_Текстовые_файлы_Операции"
Option Explicit    'Нам чужого не надо, но своё мы возьмем, чьё бы оно ни было...

Sub Добавить_B_текстовый_файл()
   Dim strFile_Path As String
   Dim username As String
       username = Environ("UserName")
       strFile_Path = "D:\OneDrive\Лог отчета по клаймам.txt"
   Open strFile_Path For Append As #1
   Print #1, " "
   Print #1, "Отчёт по клаймам заполнил " & username & " " & Now   '
   Close #1
End Sub

Sub Чтение_текстового_файла_в_переменную()
    Dim strFileName As String
    strFileName = "D:\OneDrive\Лог отчета по клаймам.txt"
    Dim strFileContent As String
    Dim iFile As Integer
    iFile = FreeFile
    Open strFileName For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    MsgBox strFileContent
    ThisWorkbook.FollowHyperlink ("D:\OneDrive\Лог отчета по клаймам.txt") 'Открытие_текстового_файла_в_блокноте
End Sub


Sub Открытие_текстового_файла_в_блокноте()
    ThisWorkbook.FollowHyperlink ("D:\OneDrive\Лог отчета по клаймам.txt") 'Открытие_текстового_файла_в_блокноте
End Sub

Sub Открыть_Лог_отчета_по_клаймам()
    ThisWorkbook.FollowHyperlink ("D:\OneDrive\Лог отчета по клаймам.txt") 'Открытие Лога отчета по клаймам.txt_в_блокноте
End Sub

Sub Открыть_Лог_отчета_по_клаймам_И() 'Вывод в  MsgBox последней строки Лога отчета по клаймам.txt
    Dim fso As Object
    Dim file As Object
    Dim TextStream As Object
     Dim str As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile("D:\OneDrive\Лог отчета по клаймам.txt")
    Set TextStream = file.OpenAsTextStream(1)
    str = vbNullString
    While Not TextStream.AtEndOfStream
        str = TextStream.ReadLine() & vbCrLf    'Str = Str & TextStream.ReadLine() & vbCrLf 'Вывод в  MsgBox всех строк
    Wend
    MsgBox str
    TextStream.Close
   
    Set fso = Nothing  ' Освобождаем объект
    Set file = Nothing  ' Освобождаем объект
End Sub
