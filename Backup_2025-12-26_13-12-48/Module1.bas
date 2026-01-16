Attribute VB_Name = "Module1"
Option Explicit
Public gRibbon As IRibbonUI
Public pressed As Boolean
Public Включение As Boolean
Public прочитанная_переменная As Boolean
Public str1 As String
Public ff As Integer
Private CheckBoxValue As Boolean

Sub RibbonLoaded(ribbon As IRibbonUI)
    Dim str1 As String
    Dim iniPath As String
    Dim fileExists As Boolean
    
    ' 1. Проверяем, что ribbon передан корректно
    If ribbon Is Nothing Then
        Debug.Print "Ошибка: ribbon = Nothing в RibbonLoaded"
        Exit Sub
    End If
    
    ' 2. Формируем путь к INI-файлу
    iniPath = Environ("APPDATA") & "\Microsoft\AddIns\1234.ini"
    
    ' Проверяем существование файла
    fileExists = (Dir(iniPath, vbNormal) <> "")
    
    ' 3. Читаем значение из INI (с обработкой ошибок)
    On Error Resume Next
    If fileExists Then
        Open iniPath For Input As #1
        Line Input #1, str1
        Close #1
    Else
        str1 = "" ' Файл не найден > используем значение по умолчанию
    End If
    Err.Clear
    On Error GoTo 0
    
    ' 4. Очищаем строку
    str1 = Replace(str1, "#", "")
    
    ' 5. Инициализируем gRibbon
    Set gRibbon = ribbon
    
    ' 6. Активируем вкладку, только если условие выполнено
    If str1 = "True" Then
        On Error Resume Next ' Защищаемся от ошибок активации
        gRibbon.ActivateTabQ "Tab4", "Надстройка2"
        If Err.Number <> 0 Then
            Debug.Print "Ошибка активации вкладки: " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
    End If
    
    Debug.Print "RibbonLoaded выполнен. gRibbon инициализирован."
End Sub
 

Sub Test(control As IRibbonControl)
   Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev"
       username = Environ("UserName")  ' Получаем имя пользователя.
       If username <> SpecifiedUserName Then: Exit Sub
    Dim Start As Date
    Dim iLastRow As Long
    Dim dokumentov  As Long
    Dim vRetVal                                          'Для получения выбранного значения.
   
'    GoTo Blok_2
 '-----------------------------------------------------------------------
    Dim sheet As Worksheet                               'Это Private Sub ПроверкаЛиста()
    Dim cell As Range
    Dim sName As String                                  'Создаём переменную, в которую поместим имя листа.
    sName = "Валидация ПД"                               'Помещаем в переменную имя листа

    On Error Resume Next
        If Worksheets(sName) Is Nothing Then             'Действия, если листа нет
        If MsgBox("На данном листе эта кнопка не работает," & _
                   vbNewLine & "создать нужный лист и построить таблицу?", vbYesNo) <> vbYes Then Exit Sub
            Worksheets.Add.Name = "Валидация ПД"
        End If
'    ActiveWindow.DisplayGridlines = False               'Отключить сетку
'    Worksheets("Валидация").UsedRange.ClearContents     'Очистить содержимое
    Range("A1:I1").Interior.Color = RGB(68, 84, 106)
    Range("A1:I1").Font.Color = RGB(255, 255, 255)
    Range("A1").Value = "№ п/п"
    Range("B1").Value = "№ Договора Займа"
    Range("C1").Value = "Количество документов"
    Range("F1").Value = "Дата"
    Range("G1").Value = "Время"
    Range("A1:Z1").WrapText = False
    Range("A1:Z1").VerticalAlignment = xlCenter          ' Выравнивание по центру
    Range("A1:Z1").HorizontalAlignment = xlLeft          ' Выравнивание по левому краю
    Range("A1:Z1").Font.Size = 9
    Range("A1:Z1").Font.Name = "Calibri"
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    Dim rng As Range:                                    'Тонкая граница вокруг всех ячеек в диапазоне
        Set rng = Range("A1:G1")
        With rng.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Range("I1:I99999").NumberFormat = "#,##0.00"
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True                      'Закрепить верхнюю строку
     iLastRow = Cells(Rows.count, 3).End(xlUp).Row
        Cells(iLastRow + 1, 3).Select
    Worksheets("Валидация").Background.Fill.ForeColor.RGB = RGB(192, 192, 192)
'Blok_2:
  '----------------------------------------------------------------------------
    vRetVal = InputBox("Укажите количество документов:", "Запрос данных", "4", 15500, 8200)
        If Val(vRetVal) = 0 Then
        MsgBox "количество документов должно быть целым числом больше нуля!", vbCritical, "Количество документов"
        Exit Sub
    End If
         
    iLastRow = Cells(Rows.count, 3).End(xlUp).Row
        Cells(iLastRow + 1, 3).Select
    
    Cells(iLastRow + 1, 3).Value = vRetVal
    GoTo Instruk
'Stop
    Sleep (300)
    SetCursorPos 1690, 1020          'клик
    
'    If MsgBox("Сохранить?", vbYesNo) <> vbYes Then Exit Sub

               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    
    Start = Timer ' Пауза для ................................
            Do While Timer < Start + 4
                DoEvents
            Loop

'If MsgBox("Копировать № Договора?", vbYesNo) <> vbYes Then Exit Sub
       
    SetCursorPos 1100, 345           'клик на поле с № Договора
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
           Sleep (300)
'If MsgBox("Дальше?", vbYesNo) <> vbYes Then Exit Sub
           Application.SendKeys ("^a")
           Sleep (500)
           Application.SendKeys ("^c")
'                Start = Timer ' Пауза для ................................
'                               Do While Timer < Start + 1
'                                   DoEvents
'                               Loop
'    AppActivate ("Валидация_My_2.xlsm")  ' Активирую книгу.
  
'    If MsgBox("Выбрать первую пустую ячейку столбца B?", vbYesNo) <> vbYes Then Exit Sub
     iLastRow = Cells(Rows.count, 2).End(xlUp).Row
        Cells(iLastRow + 1, 2).Select

'    If MsgBox("Вставить?", vbYesNo) = vbNo Then Exit Sub
    Sleep (300)
    ActiveSheet.Paste
Instruk:
    Cells(ActiveCell.Row, 6).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 6).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 7).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 7).NumberFormat = "hh:mm:ss"
    Cells(ActiveCell.Row, 8).FormulaLocal = "=СУММЕСЛИ(F2:F5000;СЕГОДНЯ();C2:C5000)"
    Cells(ActiveCell.Row, 8).Value = Cells(ActiveCell.Row, 8).Value
'    Cells(ActiveCell.Row, 9).FormulaLocal = "=H1/442*2800"
    Cells(ActiveCell.Row, 9).Value = Cells(ActiveCell.Row, 8) / 442 * 2800
    ActiveSheet.Range("I1") = Cells(ActiveCell.Row, 9).Value
    Cells(ActiveCell.Row, 9).Value = Cells(ActiveCell.Row, 9).Value
    ActiveCell.Offset(0, -2).FormulaLocal = "=СЧЁТЕСЛИ(F2:F5000;СЕГОДНЯ())"
    ActiveCell.Offset(0, -2).Value = ActiveCell.Offset(0, -2).Value
   
     Worksheets("Валидация ПД").Columns("A:Z").AutoFit
    SendKeys "{NUMLOCK}"
'    Sleep (300)

        With ActiveSheet.Range("H1")
            .Formula = "=SUMIFS(C2:C50000, F2:F50000, TODAY())"
            .Value = .Value
        End With
 
    dokumentov = ActiveSheet.Range("H1")
 
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "Бокс_1" ' Документов обновится при выполнении этой процедуры
    End If
    
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "Бокс_2" ' Сегодня
    End If
    
     If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "Бокс_3" ' Сумма
     End If

        If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "Бокс_Градусы"
    End If
'
     If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "editBox_Dollar"
    End If
    
     If Not gRibbon Is Nothing Then
       gRibbon.InvalidateControl "editBox_Строк" ' "строк" обновится при выполнении этой процедуры
    End If
    
    If Not gRibbon Is Nothing Then
       gRibbon.InvalidateControl "editBox_папок" ' "папок" обновится при выполнении этой процедуры
    End If
    
    If Not gRibbon Is Nothing Then
       gRibbon.InvalidateControl "editBox_файлов" ' "файлов" обновится при выполнении этой процедуры
    End If
    
    
    
    
'    Sleep (300)
    
'    MsgBox "Сработала процедура, заданная в onAction элемента " & control.ID

'Call Три_Документа

'------------------------------------------------------------------------------------
'    If MsgBox("Запустить Кликера?", vbYesNo) <> vbYes Then Exit Sub
'    SetCursorPos 1955, 980          'клик
'    mouse_event &H2, 0, 0, 0, 0
'    Sleep (300)
'    mouse_event &H4, 0, 0, 0, 0
'----------------------------------------------------------------------------------
    
   Dim r&, r1&, Rn&, C&, C1&, Cn&
    Set rng = Range("A2:I77")
    r1 = rng.Row: Rn = rng.Rows.count + r1 - 1
    C1 = rng.Column: Cn = rng.Columns.count + C1 - 1
    
    For r = r1 To Rn
        ' Проверяем, есть ли данные в строке перед установкой обрамления
        If Application.WorksheetFunction.CountA(Rows(r)) > 0 Then
            ' Устанавливаем обрамление для строк с данными
            With Intersect(Rows(r), rng)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With
            
            ' Устанавливаем заливку только для нечетных строк
            If r Mod 2 <> 0 Then
                With Intersect(Rows(r), rng)
                    .Interior.Color = RGB(232, 232, 232)
                End With
            End If
        End If
    Next r
       
End Sub

Sub ToggleButton1_OnAction(control As IRibbonControl, pressed As Boolean)
    If pressed Then
        Включение = True
    Else
        Включение = False
    End If
        ff = FreeFile
        Open ThisWorkbook.Path & "\1234.ini" For Output As ff
        Write #ff, Включение
        Close ff
End Sub

Public Sub CaptureState(CheckBox As IRibbonControl, ByRef ReturnValue)
    CheckBoxValue = ReturnValue
End Sub

Public Sub GetPressed(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "toggleButton_236"
            returnedVal = прочитанная_переменная
    End Select
End Sub


Sub CallTest() ' Для горячей клавиши
    Call Test(Nothing)
End Sub
 
Sub CallОткрытьПапкуДинамика() ' Для горячей клавиши
    Call ОткрытьПапкуДинамика(Nothing)
End Sub

Sub CallОткрытьФайлДинамика() ' Для горячей клавиши
    Call ОткрытьФайлДинамика(Nothing)
End Sub

Sub CallTVozduha() ' Для загрузки при открытии книги
    Call TVozduha(Nothing)
End Sub
 
Sub dokumentov(editBox As IRibbonControl, ByRef text)
On Error GoTo Instruk
    Dim dokumentov  As Long
    dokumentov = ActiveSheet.Range("H1")
    text = "   " & dokumentov
Instruk:
    Exit Sub
End Sub

Sub Segodnja(editBox As IRibbonControl, ByRef text)
    Dim dtToday As Date
        dtToday = Now
    text = "   " & Format(dtToday, "dd mmmm yyyy г.")
End Sub

Private Sub Summa(editBox As IRibbonControl, ByRef text)
On Error GoTo Instruk
    Dim Summa  As Long
    Summa = ActiveSheet.Range("I1")
    text = "   " & Format(Summa, "0.00" & " руб")
Instruk:
    Exit Sub
End Sub

Private Sub ОткрытьФайлДинамика(control As IRibbonControl)
'    Workbooks.Open FileName:="C:\Users\Хозяин\Desktop\Динамика 2025 Электрозаводская.xlsx"
    Workbooks.Open FileName:="Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx"
End Sub




