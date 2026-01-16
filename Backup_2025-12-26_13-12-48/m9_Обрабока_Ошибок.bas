Attribute VB_Name = "m9_Обрабока_Ошибок"
'1. Вывод статистики в окно Immediate:
'            Debug.Print "Пользователь: " & username & "  |" & "
'    On Error Resume Next

Sub А2345(Optional Dummy)
    MsgBox "Не определён Знак", vbExclamation, "Ошибка выполнения": Exit Sub
End Sub


Sub ЗахватТекста(Optional Dummy) 'As Long 'для числового значения и As String для текстового
On Error GoTo Instruk
   Dim EditBoxТекст  As Long 'для числового значения и As String для текстового
   EditBoxТекст = text
   ActiveSheet.Range("AT1") = EditBoxТекст
   Exit Sub
Instruk:
 MsgBox "Ошибка: " & Err.Description
    Exit Sub
End Sub




Sub Обрабока_Ошибок_норм(Optional Dummy)

    Dim ИмяПроцедуры As String
    Dim ИмяМодуля As String
    ИмяМодуля = ActiveWorkbook.VBProject.VBE.ActiveCodePane.CodeModule.Parent.Name
    ИмяПроцедуры = "ВСТАВИТЬ_ИМЯ(control As IRibbonControl)"
    
    On Error GoTo ErrHandler
    ' остальной код
    Exit Sub
ErrHandler:     'MsgBox "ТЕКСТ"
     MsgBox "Ошибка: " & Err.Description & "https://my-calend.ru/goroskop/telec/ " & _
     vbNewLine & "Имя Модуля: " & Application.VBE.ActiveCodePane.CodeModule.Name & _
     vbNewLine & "Имя Процедуры: " & ИмяПроцедуры, vbExclamation, "Ошибка выполнения"
'    Range("N1").Value = Err.Description
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
   .SetText Err.Description: .PutInClipboard
End With
Dim URL As String
    URL = "https://yandex.ru/search/?text=vba+excel....&lr=65&clid=2411726%2F/"
    ShellExecute 0, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus
    On Error GoTo 0  ' Сброс обработчика ошибок
    Exit Sub
End Sub


'Сообщение ErrHandler: MsgBox в VBA отображается не только при ошибке, но и при её отсутствии.
'Это происходит из-за того, что код в ErrHandler выполняется при достижении этой точки в программе, даже если ошибок не возникло.

'Чтобы решить проблему, нужно добавить оператор Exit Sub перед ErrHandler.
'Это позволит завершить выполнение функции перед запуском кода в ErrHandler

'В этом случае сообщение ErrHandler: MsgBox будет отображаться только при ошибке,
'а при её отсутствии программа завершится без выполнения кода в ErrHandler

'----------------------------------------------------------------------------------------------------------------------------
'СИНТАКСИС ВЫРАЖЕНИЙ С On Error:
1 On Error GoTo Stroka

'Включает алгоритм обнаружения ошибок и, в случае возникновения ошибки, передает управление операторам обработчика ошибок с указанной в выражении строки. Stroka – это метка, после которой расположены операторы обработчика ошибок.
1 On Error Resume Next

'Включает алгоритм обнаружения ошибок и, в случае возникновения ошибки, передает управление оператору, следующему за оператором, вызвавшем ошибку.
1 On Error GoTo 0

'Отключает любой включенный обработчик ошибок в текущей процедуре.
