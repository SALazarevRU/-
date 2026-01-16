Attribute VB_Name = "m9_Проверка_Наличия_Библиотек"
Sub Проверка_Наличия_Библиотек()

    Dim ИмяПроцедуры As String
    Dim ИмяМодуля As String
    ИмяМодуля = ActiveWorkbook.VBProject.VBE.ActiveCodePane.CodeModule.Parent.Name
    ИмяПроцедуры = "ВСТАВИТЬ_ИМЯ(control As IRibbonControl)"
    On Error GoTo ErrHandler
 
Dim i As Integer
Dim currentDateTime As Date
    currentDateTime = Now
    Debug.Print "Текущая Дата и Время: "; currentDateTime
With ThisWorkbook.VBProject.References
  For i = 1 To .Count
  
    Debug.Print .Item(i).GUID, .Item(i).Description, .Item(i).Major, .Item(i).Minor
    If .Item(i).GUID = "{420B2830-E718-11CF-893D-00A0C9054228}" Then
      Exit Sub
    End If
  Next i
  'Microsoft scripting
  .AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
End With

 Exit Sub
ErrHandler:     MsgBox "ТЕКСТ"
     MsgBox "Ошибка: " & Err.Description & "ИСКОМОЕ: " & _
     vbNewLine & "Имя Модуля: " & ИмяМодуля & _
     vbNewLine & "Имя Процедуры: " & ИмяПроцедуры, vbExclamation, "Ошибка выполнения"
    
    On Error GoTo 0  ' Сброс обработчика ошибок
    Exit Sub

End Sub
