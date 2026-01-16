Attribute VB_Name = "m9_Загрузка_Списка_Подпапок"
Option Explicit
Option Compare Text
 
Private Sub ПроверкаЛиста()
   Dim sheet As Worksheet
   Dim cell As Range
   Dim sName As String 'Создаём переменную, в которую поместим имя листа.
   sName = "Валидация" 'Помещаем в переменную имя листа
   
   Application.EnableEvents = False 'Отключаем отслеживание событий
   
   On Error Resume Next
   If Worksheets(sName) Is Nothing Then  'действия, если листа нет
       Worksheets.Add.Name = "Валидация"
   End If
   Worksheets("Валидация").UsedRange.ClearContents
   
  Application.EnableEvents = True
End Sub
 
 
Sub ЗагрузкаСпискаПодпапок() ' процедура ЗагрузкаСпискаПодпапок на лист "Валидация" если его нет- создается
    Dim i           As Long
    On Error GoTo ErrHandler
 
    Dim L           As String
'    L = "C:\Users\Хозяин\Desktop\Сканы АБВ"
    L = "C:\Users\s.lazarev\Desktop\Сканы АБВ"
    Dim coll        As Collection
    Set coll = SubFoldersCollection(L)
    
    
'----------------------------------------------------------------------
'Это Private Sub ПроверкаЛиста()
   Dim sheet As Worksheet
   Dim cell As Range
   Dim sName As String 'Создаём переменную, в которую поместим имя листа.
   sName = "Валидация" 'Помещаем в переменную имя листа
   
   Application.EnableEvents = False 'Отключаем отслеживание событий
   
   On Error Resume Next
   
   If Worksheets(sName) Is Nothing Then  'действия, если листа нет
       Worksheets.Add.Name = "Валидация"
   End If
   Worksheets("Валидация").UsedRange.ClearContents
   
  Application.EnableEvents = True
'--------------------------------------------------------------------------
    
 
    With ActiveWorkbook.Worksheets("Валидация")    ' замените на конкретное имя вашего листа
 
        Dim nextRow As Long
        nextRow = .Cells(.Rows.Count, "C").End(xlUp).Row + 1
        If .Cells(1, "C").Value = "" Then nextRow = 1
 
        For i = 1 To coll.Count
            .Cells(nextRow, 3).Value = coll(i)
            nextRow = nextRow + 1
        Next i
 
    End With
 
    Exit Sub
ErrHandler:
    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Ошибка выполнения"
   On Error GoTo 0
   Exit Sub
End Sub
 
Function SubFoldersCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "*") As Collection
    Dim curfold     As Object
    Dim Folder      As Object
    On Error GoTo ErrHandler
 
    Set SubFoldersCollection = New Collection
 
    Dim FSO         As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
 
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    Set curfold = FSO.GetFolder(FolderPath)
 
    If Not curfold Is Nothing Then
 
        For Each Folder In curfold.Subfolders
 
            If Folder.Path Like FolderPath & Mask Then
 
                ' только имена подпапок
                SubFoldersCollection.Add Folder.Name
 
                '                ' полные пути к подпапкам с именами подпапок
                '                SubFoldersCollection.Add folder.Path & "\"
            End If
 
        Next Folder
 
    End If
 
    Set FSO = Nothing
    Set curfold = Nothing
    Exit Function
 
ErrHandler:
    MsgBox "Ошибка при получении подпапок: " & Err.Description, vbCritical, "Ошибка"
    Set SubFoldersCollection = Nothing
    Set FSO = Nothing
    Set curfold = Nothing
End Function

