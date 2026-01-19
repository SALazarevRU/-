Attribute VB_Name = "Module6"
Option Explicit

Sub Case_Example1()
    Dim ФИО As String
    ФИО = user_name
    Select Case Cell_Value
        Case "v.petrov"
             ФИО = "Вася Петров"
        Case "m.ivanova”"
             ФИО = "Маша Иванова"
        Case "g.sidorov"
             ФИО = "Гоша Сидоров"
    End Select
End Sub


'Чтобы установить автофильтр в VBA Excel на столбец, используя переменную "Сегодня()",
'вам нужно использовать функцию  Date  для получения текущей даты. Вот пример кода:

 
Sub AutoFilterByToday()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filterColumn As Long
    Dim todayDate As Date

    ' Укажите лист, с которым работаете
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Замените "Sheet1" на имя вашего листа

    ' Укажите номер столбца для фильтрации (например, 1 для столбца A)
    filterColumn = 1

    ' Получаем последнюю строку с данными в столбце
    lastRow = ws.Cells(Rows.Count, filterColumn).End(xlUp).Row

    ' Получаем текущую дату
    todayDate = Date

    ' Применяем автофильтр
    With ws.Range(ws.Cells(1, filterColumn), ws.Cells(lastRow, filterColumn))
        .AutoFilter Field:=1, Criteria1:=todayDate
    End With

End Sub
 

Пояснения:

 Dim ws As Worksheet : Объявляет переменную  ws  как объект Worksheet.
 Dim lastRow As Long : Объявляет переменную  lastRow  для хранения номера последней строки с данными.
 Dim filterColumn As Long : Объявляет переменную  filterColumn  для хранения номера столбца, который нужно фильтровать.
 Dim todayDate As Date : Объявляет переменную  todayDate  для хранения текущей даты.
 Set ws = ThisWorkbook.Sheets("Sheet1") : Устанавливает, с каким листом вы работаете. Важно: Замените  "Sheet1"  на имя вашего листа.
 filterColumn = 1 : Устанавливает номер столбца для фильтрации. В данном случае, это столбец A (первый столбец).
 lastRow = ws.Cells(Rows.Count, filterColumn).End(xlUp).Row : Определяет последнюю строку с данными в указанном столбце.
 todayDate = Date : Получает текущую дату.
 With ws.Range(ws.Cells(1, filterColumn), ws.Cells(lastRow, filterColumn)) : Определяет диапазон для фильтрации, начиная с первой строки и заканчивая последней строкой в указанном столбце.
 .AutoFilter Field:=1, Criteria1:=todayDate : Применяет автофильтр к указанному диапазону, фильтруя по текущей дате.

Как использовать этот код:

Откройте Excel.
Нажмите  Alt + F11 , чтобы открыть редактор VBA.
Вставьте этот код в модуль (Insert -> Module).
Измените  "Sheet1"  на имя вашего листа.
Измените  filterColumn = 1 , если вам нужно фильтровать другой столбец.
Запустите макрос (нажмите  F5  или кнопку "Run").

Важно:

Убедитесь, что формат даты в вашем столбце соответствует формату, который возвращает функция  Date . Если у вас другой формат даты, возможно, потребуется преобразовать дату в нужный формат перед применением фильтра.
Если автофильтр уже включен, этот код его выключит и включит снова с новыми критериями.

