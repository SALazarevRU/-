Attribute VB_Name = "m9_Проверка_наличия_Листа"
Private Sub ПроверкаЛиста(Optional Dummy)

   Dim sName As String 'Создаём переменную, в которую поместим имя листа.
   sName = "Валидация" 'Помещаем в переменную имя листа
   
'   MsgBox TypeName(sName)  ' проверить тип переменной можно с помощью функции TypeName
   
   Application.EnableEvents = False 'Отключаем отслеживание событий
   
   On Error Resume Next
   If Workbooks("Лист Microsoft Excel (3).xlsx").Worksheets(sName) Is Nothing Then  'действия, если листа нет
       Worksheets.Add.Name = "жопа"
   End If
   Workbooks("Лист Microsoft Excel (3).xlsx").Worksheets("sName").UsedRange.ClearContents
   
  Application.EnableEvents = True
End Sub
