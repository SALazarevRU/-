Attribute VB_Name = "m9_Оптимизация_Кода"
Sub ОптимизацияКода(Optional Dummy)
 ActiveWorkbook.Sheets("Гороскоп").Delete
' ThisWorkbook.
Application.DisplayAlerts = False  ' Отключить диалоговое окно отображение предупреждений и сообщений

Application.ScreenUpdating = False  'отключаем обновление экрана

Application.Calculation = xlCalculationManual 'Отключаем автопересчет формул

Application.EnableEvents = False 'Отключаем отслеживание событий

ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False 'Отключаем разбиение на печатные страницы
 
'Непосредственно код заполнения ячеек
Dim lr As Long
For lr = 1 To 10000
    Cells(lr, 1).Value = lr 'для примера просто пронумеруем строки
Next
 
Application.ScreenUpdating = True 'Возвращаем обновление экрана

Application.Calculation = xlCalculationAutomatic 'Возвращаем автопересчет формул

Application.EnableEvents = True 'Включаем отслеживание событий

Application.DisplayAlerts = True ' Снова включить отображение предупреждений и сообщений
End Sub

'Модификатор специальных возможностей Private позволяет отображать подпроцедуру во все остальные подпроцедуры,
'но только с модулем, в котором он находится.
'
'Кроме того, как вы видели в предыдущем разделе, частные подпрограммы не отображаются в конечным пользователем
'при нажатии на кнопку «Макросы» в книге.

Private Sub Задержка()

    Start = Timer
    Do While Timer < Start + 0.5 '0.5 = полсекунды
        DoEvents
    Loop
    
End Sub
