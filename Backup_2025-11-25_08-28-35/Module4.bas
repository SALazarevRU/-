Attribute VB_Name = "Module4"

'Макрос, который записывает значения двух переменных в свойства указанной книги, может выглядеть так:
Sub ЗаписатьСвойстваКниги()
    ' Объявляем переменные как объекты типа Workbook:
    Dim oWorkbook As Workbook
    Dim sFilePath As String
    ' Указываем путь к файлу для переменной sFilePath:
    sFilePath = "D:\VBAF1\VBA Functions.xlsm"
    ' Присваиваем объект переменной oWorkbook с помощью оператора `Set`:
    Set oWorkbook = Workbooks.Open(sFilePath)
    ' Используем переменную oWorkbook для записи значений в свойства книги:
    oWorkbook.Name = "Книга1"
    oWorkbook.title = "Книга2"
End Sub

Sub Записать_значения_в_свойства_книги()
    
    Dim a As Integer
    Dim b As Integer
    
    Dim жопа As Variant
    
    Workbooks("Валидация_My_2.xlam").Properties.Item("Значение переменной жопа").Value = грязная
    Workbooks("Валидация_My_2.xlam").Item("Значение переменной b").Value = b
End Sub

Sub ПримерИспользованияПользовательскихСвойствКнигиExcel_МОЙ()
'    DDocALL ActiveWorkbook    ' удаляем все ранее назначенные пользовательские свойства
    ' и записываем новые:
    SDoc ActiveWorkbook, "ICQ", "58-36-318"
    SDoc ActiveWorkbook, "Skype", "ExcelVBA.ru"
    SDoc ActiveWorkbook, "Сайт", "http://ExcelVBA.ru/"
 
    ' теперь можно закрыть файл, предварительно его сохранив
    ' а потом, после очередного открытия, считать сохранённые свойства:

    txt = GDoc(ActiveWorkbook, "ICQ") & vbNewLine & GDoc(ActiveWorkbook, "Сайт")
    MsgBox txt, vbInformation, "Пользовательские свойства книги Excel"
    ' и удалить ненужные
    DDoc ActiveWorkbook, "ICQ"
End Sub

Sub ПримерИспользованияПользовательскихСвойствКнигиExcel()
'    DDocALL ActiveWorkbook    ' удаляем все ранее назначенные пользовательские свойства
'    ' и записываем новые:
'    SDoc ActiveWorkbook, "ICQ", "58-36-318"
'    SDoc ActiveWorkbook, "Skype", "ExcelVBA.ru"
'    SDoc ActiveWorkbook, "Сайт", "http://ExcelVBA.ru/"
 
    ' теперь можно закрыть файл, предварительно его сохранив
    ' а потом, после очередного открытия, считать сохранённые свойства:

    txt = GDoc(ActiveWorkbook, "ICQ") & vbNewLine & GDoc(ActiveWorkbook, "Сайт")
    MsgBox txt, vbInformation, "Пользовательские свойства книги Excel"
    ' и удалить ненужные
'    DDoc ActiveWorkbook, "ICQ"
End Sub


'With Workbooks("Книга")
'    .Properties.Item("Значение переменной a").Value = a
'    .Properties.Item("Значение переменной b").Value = b
'End With

Sub Show_CustomDocumentProperties() 'просмотреть все пользовательские свойства, сохранённые в файле, - то используйте это
    ' выводит список всех пользовательских свойств в книге, из которой запускается макрос
    If ThisWorkbook.CustomDocumentProperties.Count > 0 Then
        For Each cdp In ThisWorkbook.CustomDocumentProperties
            txt = txt & cdp.Name & ":" & vbTab & cdp.Value & vbNewLine
        Next
        MsgBox txt, vbInformation, "Список пользовательских свойств в книге"
    End If
End Sub


Sub SDoc(ByRef wb As Workbook, ByVal VarName As String, ByVal VarValue As Variant)
    ' сохранение пользовательского свойства в книге Excel
    DDoc wb, VarName    ' удаляем свойство, если оно уже есть
    ' и создаём новое с нужным значением
    wb.CustomDocumentProperties.Add VarName, False, msoPropertyTypeString, CStr(VarValue)
End Sub
 
Sub DDoc(ByRef wb As Workbook, ByVal VarName As String)
    ' удаление пользовательского свойства из книги Excel
    If wb.CustomDocumentProperties.Count > 0 Then    ' если они вообще есть
        For Each cdp In wb.CustomDocumentProperties    ' перебираем все свойства
            If cdp.Name = VarName Then cdp.Delete: Exit Sub    ' удаляем
        Next
    End If
End Sub
 
Sub DDocALL(ByRef wb As Workbook)
    ' удаление ВСЕХ пользовательских свойств из книги Excel
    If wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In wb.CustomDocumentProperties
            cdp.Delete    ' удаляем очередное свойство
        Next
    End If
End Sub
 
Function GDoc(ByRef wb As Workbook, ByVal VarName As String) As String
    ' чтение переменной из книги Excel
    ' функция возвращает значение пользовательского свойства VarName
    ' (если нужное пользовательское свойство отсутствует, возвращает пустую строку)
    If wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In wb.CustomDocumentProperties
            If cdp.Name = VarName Then GDoc = cdp.Value
        Next
    End If
End Function
 
Function GDocB(ByRef wb As Workbook, ByVal VarName As String) As Boolean
    ' чтение переменной из книги Excel
    ' функция возвращает ПРЕОБРАЗОВАННОЕ К ТИПУ BOOLEAN значение пользовательского свойства VarName
    ' (если нужное пользовательское свойство отсутствует, возвращает FALSE)
    On Error Resume Next
    If wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In wb.CustomDocumentProperties
            If cdp.Name = VarName Then GDocB = CBool(cdp.Value)
        Next
    End If
End Function
