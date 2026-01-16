Attribute VB_Name = "m9_Проверка_Почты"
Private ПроверкаNewПисем As String
Private Sub ПроверкаПочты(control As IRibbonControl)
'Sub ПроверкаПочты()
    Dim sizesOSTold As Long
    Dim sizesOSTnew  As Long
'    Dim MyCollection As Collection
'    Set MyCollection = New Collection
'    Dim FN1 As String, FN2 As String
    Dim sizes(1 To 2) As Double, dif As Double
    
    Call ЗапуститьOutlookСвёрнуто
   ' Ждем, пока клиент откроется
    Application.Wait Now + TimeValue("00:00:08")
    sizesOSTold = ThisWorkbook.CustomDocumentProperties.Item("BookSetting").Value

    AppActivate ("Лист Microsoft Excel (3).xlsx - Excel") 'вытягиваем на первый план
    MsgBox "Old размер файла .ost : " & sizesOSTold
    Debug.Print sizesOSTold
'            FN1 = "C:\Users\Хозяин\AppData\Local\Microsoft\Outlook\s.a.lazarev@yandex.ru.ost"
    Call CloseOutlook
     AppActivate ("Лист Microsoft Excel (3).xlsx - Excel") 'вытягиваем на первый план
   ' Ждем, пока клиент закроется
    Application.Wait Now + TimeValue("00:00:02")
    Dim FSO, Folder, Size
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder("C:\Users\s.lazarev\AppData\Local\Microsoft\Outlook")
    sizesOSTnew = Folder.Size
    
'            sizes_1 = FileLen(FN1) ' Запись размеров файлов в байтах в массив'
'            sizesOSTnew = sizes_1

    MsgBox "New размер файла .ost : " & sizesOSTnew 'MsgBox для отладки
    Debug.Print sizesOSTnew
    MsgBox "Размеры файлов .ost : " & vbNewLine & sizesOSTold & vbNewLine & sizesOSTnew 'MsgBox для отладки
          If sizesOSTnew > sizesOSTold Then
          Debug.Print sizesOSTnew - sizesOSTold
              ПроверкаNewПисем = "Прибытие" 'Вывод в editBox Ribbon
              MsgBox "Прибытие"
          Else
              ПроверкаNewПисем = "Нет новых"
          End If
'   Вывод результата в editBox:
    If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBox_Почта"
    End If
'   сохранить в CustomDocumentProperties текущией книги значение sizesOSTnew
        With Workbooks("Надстройка2.xlam").CustomDocumentProperties
            On Error Resume Next
                .Add Name:="BookSetting", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=""
                .Item("BookSetting").Value = sizesOSTnew
        End With
'    Set Folder = Nothing
End Sub

Private Sub Сообщение(editBox As IRibbonControl, ByRef Text)
    Text = "   " & ПроверкаNewПисем
End Sub



'РАБОТАЕТ===================================================================================================================
Private Sub Save_Val_in_DocProp() 'сохранить в CustomDocumentProperties текущией книги значение
    Dim sMyVal$: sMyVal = sizesOSTnew ' значение, которое нужно сохранить
    With Workbooks("Надстройка2.xlam").CustomDocumentProperties
        On Error Resume Next
        .Add Name:="BookSetting", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=""
        .Item("BookSetting").Value = sMyVal
    End With
End Sub
Private Sub Restore_Val_from_DocProp()    ' считать из CustomDocumentProperties текущией книги сохранённое значение
  Dim sMyVal$
  On Error Resume Next
  sMyVal = ThisWorkbook.CustomDocumentProperties.Item("BookSetting").Value
   MsgBox "Что было записано : " & sMyVal
  If Err Then sMyVal = "Error"
End Sub



