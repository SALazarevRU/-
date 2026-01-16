Attribute VB_Name = "m9_Поиск_Ямайка"
' Module    : m2_Поиск_New
' Purpose   : Выделить ячейку в диапазоне (B3:B13) ПРИ УСЛОВИИ Offset(, 2) = "1" введя в InputBox несколько символов из фио.  В коде обработать
'             три исключения: "Поле ввода осталось пустым", "Значение не найдено" и "Найдено несколько ячеек с введёнными символами"
' Status   :  Работает... Но БЕЗ УСЛОВИЯ.( + если ввести сразу полное значение, то выделит со второго раза, с Оk "Найдено несколько значений.")
'------------------------------------------------------------------------------------------------------------------------

'Оставьте поле ввода пустым:     .Ожидаемый результат: остановка процедуры. Результат: остановка процедуры.
'Введите в поле ввода: иванов    .Ожидаемый результат: Иванов Олег. Результат: Несколько значений - > Запрос бОльшего кол-ва символов
'Введите в поле ввода: иванов о  .Ожидаемый результат: Иванов Олег. Результат: Иванов Олег. Отлично, End Sub.
'Введите в поле ввода: иванов олег  .Ожидаемый результат: Иванов Олег. Результат: Несколько значений - > Запрос бОльшего кол-ва символов

Public Sub Поиск_Ямайка(control As IRibbonControl)
   Call Call_Поиск_Ямайка
End Sub

Public Sub Call_Поиск_Ямайка()
    Const lSeconds As Long = 4
    Dim RezPoiska As Range
    Dim firstAddress As String
    Static Iskomoe As String
    
    ' 1. Запрашиваем искомое значение
    Iskomoe = InputBox("Введите шаблон искомого слова (символы подстановки *, ?)", , Iskomoe)
    Debug.Print "Первоначальный вариант Искомого в InputBox = " & Iskomoe
    
    If Iskomoe = "" Then Exit Sub  ' Выход, если пусто или Cancel
    
    
    ' 2. Ищем первое вхождение (в активном листе!)
    Set RezPoiska = ActiveSheet.Cells.Find( _
        What:=Iskomoe, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False _
    )
    
    ' 3. Если не найдено — сообщение и выход
    If RezPoiska Is Nothing Then
        MessageBoxTimeOut 0, "Не найдено." & _
            vbNewLine & "Открываю файл ""Реестр2"" для уточнения" & _
            vbNewLine & "Это окно закроется автоматически через 4 секунды", _
            "РЕЗУЛЬТАТ ПОИСКА", _
            vbInformation + vbOKOnly, 0&, lSeconds * 1000
        ' Call Доп_Поиск_клиента_по_ID_в_Реестре2  ' Раскомментируйте при необходимости
        Exit Sub
    End If
    
    ' 4. Сохраняем адрес первого найденного
    firstAddress = RezPoiska.Address
    RezPoiska.Select  ' Выделяем первое совпадение
    
    
    ' 5. Создаём экземпляр clsFindBar и ищем все совпадения
    With New clsFindBar
        Do
            Set RezPoiska = ActiveSheet.Cells.FindNext(RezPoiska)
            If RezPoiska Is Nothing Then Exit Do  ' Защита от Nothing
            .Add RezPoiska
            If RezPoiska.Address = firstAddress Then Exit Do  ' Вернулись к первому
        Loop
        
        ' 6. Показываем результаты, если их больше одного
        If .Count > 1 Then
            .Show
        Else
            ' Если только одно — явно выделяем
            ActiveSheet.Range(firstAddress).Select
        End If
    End With
End Sub


