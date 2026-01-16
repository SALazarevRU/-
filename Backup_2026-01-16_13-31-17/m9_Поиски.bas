Attribute VB_Name = "m9_Поиски"
Option Explicit
Public Isk0 As Variant
Sub Поиск_от_Angry_Old_Man()
    Dim BegDann As Range: Set BegDann = Range("C3")
    Dim BegCond As Range: Set BegCond = Range("AY3")
    
    Dim Rbegin, Cbegin, Rend, DannIn, CondIn
'    Dim Isk0, Isk, i, iL, iU, n, Out, ii
    Dim Isk, i, iL, iU, n, Out, ii
    Dim Reg
    
    Dim Lfirst: Lfirst = True
    Do
        Isk0 = InputBox("Введите шаблон искомого слова ", "Поиск от Ангри Олдмана", Isk0, 12000, 6000)
        If Isk0 = "" Then
            i = MsgBox("Шаблон не введен" & vbCr & "Повторить ввод?", 33, "Шаблон не введен")
            If i = 2 Then Exit Sub '    Do
        End If
        If Lfirst Then
            Lfirst = False
            Set Reg = CreateObject("VBScript.RegExp")
            Rbegin = BegDann.Row: Rend = Split(ActiveSheet.UsedRange.Address, "$")(4)
            Cbegin = Split(BegDann.Address, "$")(1)
            DannIn = BegDann.Resize(Rend - Rbegin + 1, 1)
            CondIn = BegCond.Resize(Rend - Rbegin + 1, 1)
        End If
        Isk = Replace(Isk0, ".", "\."): Isk = Replace(Isk, "*", ".*"): Isk = Replace(Isk, "?", ".?")
        Reg.Pattern = "^" & Isk
        Reg.IgnoreCase = True       'False
        iL = LBound(DannIn, 1): iU = UBound(DannIn, 1)
        ReDim Found(iL To iU)
        
        n = 0
        For i = iL To iU
            Found(i) = (CondIn(i, iL) = 1)
            If Found(i) Then
                Found(i) = Reg.Test(DannIn(i, iL))
                If Found(i) Then n = n + 1
            End If
        Next
        
        If n = 0 Then
            i = MsgBox("Поиск по шаблону " & vbCr & vbCr & Isk0 & vbCr & vbCr & "неуспешен" & vbCr & vbCr & "Повторить ввод?", 33, "Поиск неуспешен")
            If i = 2 Then Exit Do
        Else
            ReDim Out(n)
            ii = -1
            For i = iL To iU
                If Found(i) Then
                    ii = ii + 1
                    Out(ii) = """" & Cbegin & (Rbegin - 1 + i) & """   " & DannIn(i, iL)
                End If
            Next
            
            If n = 1 Then
                Range(Replace(Split(Out(0), " ")(0), """", "")).Select
                Exit Do     ''''''''''''''''
            Else
                For i = 1 To n
                    Range(Replace(Split(Out(i - 1), " ")(0), """", "")).Activate
                    ii = MsgBox("Выбрать значение " & i & " из " & n & vbCr & vbCr & Out(i - 1), 35, "Найдено " & n & " совпадений " & """" & Isk0 & """")
                    If ii = 6 Then Exit Do  'For    ''''''''''''''''
                    If ii = 2 Then
                        Range("A1").Select
                        Exit For
                    End If
                Next
            End If
        End If
    Loop
End Sub

Sub Поиск_от_Angry_Old_Man_БЕЗУСЛОВНЫЙ_исправленный()

    Dim title As String
    title = "Поиск в столбце C (начиная с C4)"
    
    
    Dim BegDann As Range: Set BegDann = Range("C4")  ' Поиск начинается с C4
    Dim Rbegin As Long, Rend As Long
    Dim DannIn As Variant
    Dim Isk0 As String, i As Long, iL As Long, iU As Long, n As Long, ii As Long
    Dim Found() As Boolean
    Dim Out() As String
    Dim Lfirst As Boolean: Lfirst = True
    Dim selectedIndex As Long

    Do
        ' 1. Ввод шаблона
        Isk0 = InputBox("Введите текст для поиска:", title)
        If Isk0 = "" Then
            i = MsgBox("Шаблон не введён." & vbCr & "Повторить ввод?", vbExclamation + vbYesNo, "Пустой ввод")
            If i = vbNo Then Exit Sub
            GoTo NextIteration
        End If

        ' 2. Инициализация (один раз)
        If Lfirst Then
            Lfirst = False
            Rbegin = BegDann.Row
            ' Определяем последнюю заполненную строку в столбце C
            Rend = Cells(Rows.Count, "C").End(xlUp).Row
            
            ' Проверяем, есть ли данные ниже C4
            If Rend < Rbegin Then
                MsgBox "В столбце C нет данных ниже C4.", vbExclamation: Exit Do
            End If
            
            ' Считываем данные из C4 до последней заполненной строки
            If Rend = Rbegin Then
                ReDim DannIn(1 To 1, 1 To 1)
                DannIn(1, 1) = BegDann.Value
            Else
                DannIn = Range(BegDann, Cells(Rend, "C")).Value
            End If
            
            iL = LBound(DannIn, 1): iU = UBound(DannIn, 1)
        End If

        ' 3. Поиск совпадений (без учёта регистра, с обрезкой пробелов)
        ReDim Found(iL To iU) As Boolean
        n = 0
        For i = iL To iU
            If Not IsEmpty(DannIn(i, 1)) Then
                If InStr(1, Trim(CStr(DannIn(i, 1))), Trim(Isk0), vbTextCompare) > 0 Then
                    Found(i) = True
                    n = n + 1
                End If
            End If
        Next i

        ' 4. Обработка результатов
        If n = 0 Then
            i = MsgBox("Не найдено: """ & Isk0 & """" & vbCr & _
                      "Повторить ввод?", vbExclamation + vbYesNo, "Нет совпадений")
            If i = vbNo Then Exit Do
        Else
            ReDim Out(1 To n) As String
            ii = 1
            For i = iL To iU
                If Found(i) Then
                    ' Формируем строку: "C7 | Иванов Иван"
                    Out(ii) = "C" & (Rbegin + i - 1) & " | " & DannIn(i, 1)
                    ii = ii + 1
                End If
            Next i

            If n = 1 Then
                Range(Out(1)).Select
                MsgBox "Переход к ячейке выполнен.", vbInformation, "Готово"
                Exit Do
            Else
                ' Показываем список и ждём выбора
                Dim msg As String
                msg = "Найдено " & n & " совпадений. Введите номер (1–" & n & "), чтобы перейти:" & vbCr & vbCr
                For i = 1 To n
                    msg = msg & i & ". " & Out(i) & vbCr
                Next i

                Isk0 = InputBox(msg, "Выбор результата", 1)
                If Isk0 = "" Then GoTo NextIteration


                If Not IsNumeric(Isk0) Then
                    MsgBox "Введите число!", vbExclamation: GoTo NextIteration
                End If


                selectedIndex = CLng(Isk0)
                If selectedIndex < 1 Or selectedIndex > n Then
                    MsgBox "Номер вне диапазона (1–" & n & ")!", vbExclamation: GoTo NextIteration
                End If


                Range(Split(Out(selectedIndex), " | ")(0)).Select
                MsgBox "Переход выполнен.", vbInformation, "Ячейка выделена"
                Exit Do
            End If
        End If

NextIteration:
    Loop
End Sub





Sub Поиск_от_Angry_Old_Man_БЕЗУСЛОВНЫЙ() ' КОСТЫЛЬ В 37 СТРОКЕ
    Dim title
     title = "Фамилия Имя Отчество должника"
    Dim BegDann As Range: Set BegDann = Range("C4")
    Dim BegCond As Range: Set BegCond = Range("AY2")
    
    Dim Rbegin, Cbegin, Rend, DannIn, CondIn
    Dim Isk0, Isk, i, iL, iU, n, Out, ii
    Dim Reg
    
    Dim Lfirst: Lfirst = True
    Do
        Isk0 = InputBox("Введите несколько знаков ", title)
        If Isk0 = "" Then
            i = MsgBox("Шаблон не введен" & vbCr & "Повторить ввод?", 33, "Шаблон не введен")
'            If i = 2 Then Exit Do
            If i = 2 Then Exit Sub 'Instruk2

        End If
        If Lfirst Then
            Lfirst = False
            Set Reg = CreateObject("VBScript.RegExp")
            Rbegin = BegDann.Row: Rend = Split(ActiveSheet.UsedRange.Address, "$")(4)
            Cbegin = Split(BegDann.Address, "$")(1)
            DannIn = BegDann.Resize(Rend - Rbegin + 1, 1)
            CondIn = BegCond.Resize(Rend - Rbegin + 1, 1)
        End If
        Isk = Replace(Isk0, ".", "\."): Isk = Replace(Isk, "*", ".*"): Isk = Replace(Isk, "?", ".?")
        Reg.Pattern = "^" & Isk
        Reg.IgnoreCase = True       'False
        iL = LBound(DannIn, 1): iU = UBound(DannIn, 1)
        ReDim Found(iL To iU)
        
        n = 0
        For i = iL To iU
'            Found(i) = (CondIn(i, iL) = 1) ' ЭТО ДЛЯ ПОИСКА С УСЛОВИЕМ
            Found(i) = True ' КОСТЫЛЬ В ЭТОЙ СТРОКЕ ДЛЯ БЕЗУСЛОВНОГО ПОИСКА
            If Found(i) Then
                Found(i) = Reg.Test(DannIn(i, iL))
                If Found(i) Then n = n + 1
            End If
        Next
        
        If n = 0 Then
            i = MsgBox("Поиск по шаблону " & vbCr & vbCr & Isk0 & vbCr & vbCr & "неуспешен" & vbCr & vbCr & "Повторить ввод?", 33, "Поиск неуспешен")
            If i = 2 Then Exit Do
        Else
            ReDim Out(n)
            ii = -1
            For i = iL To iU
                If Found(i) Then
                    ii = ii + 1
                    Out(ii) = """" & Cbegin & (Rbegin - 1 + i) & """   " & DannIn(i, iL)
                End If
            Next
            
           If n = 1 Then
                Range(Replace(Split(Out(0), " ")(0), """", "")).Select
                Exit Do     ''''''''''''''''
            Else
                For i = 1 To n
                    Range(Replace(Split(Out(i - 1), " ")(0), """", "")).Activate
                    ii = MsgBox("Выбрать значение " & i & " из " & n & vbCr & vbCr & Out(i - 1), 35, "Найдено " & n & " совпадений " & """" & Isk0 & """")
                    If ii = 6 Then Exit Do  'For    ''''''''''''''''
                    If ii = 2 Then
                        Range("A1").Select
                        Exit For
                    End If
                Next
            End If
        End If
    Loop
End Sub


Sub Поиск_от_Angry_Old_Man_ФКБ_РАБОЧИЙ_НА_12_01_2026()
    Dim BegDann As Range: Set BegDann = Range("C3")
    Dim BegCond As Range: Set BegCond = Range("AY3")
    Dim Found() As Boolean 'Явно объявил типы переменных
    Dim cellValue As Variant 'Явно объявил типы переменных

    Dim Rbegin, Cbegin, Rend, DannIn, CondIn
'    Dim Isk0, Isk, i, iL, iU, n, Out, ii
    Dim Isk, i, iL, iU, n, Out, ii
    Dim Reg
    
Call PlayWavAPI_2

    Dim Lfirst: Lfirst = True
    Do
        Isk0 = InputBox("Введите шаблон искомого слова ", , Isk0, 12000, 6000)
        If Isk0 = "" Then
            i = MsgBox("Шаблон не введен" & vbCr & "Повторить ввод?", 33, "Шаблон не введен")
            If i = 2 Then Exit Sub '    Do
        End If
' Заменил блок с ошибкой
        If Lfirst Then
    Lfirst = False
    Set Reg = CreateObject("VBScript.RegExp")
    Rbegin = BegDann.Row: Rend = Split(ActiveSheet.UsedRange.Address, "$")(4)
    Cbegin = Split(BegDann.Address, "$")(1)
    DannIn = BegDann.Resize(Rend - Rbegin + 1, 1)
    CondIn = BegCond.Resize(Rend - Rbegin + 1, 1)
    
    ' Проверка на пустоту
    If IsEmpty(CondIn) Then
        MsgBox "Диапазон условий пуст!", vbExclamation
        Exit Sub
    End If
End If

Isk = Replace(Isk0, ".", "\."): Isk = Replace(Isk, "*", ".*"): Isk = Replace(Isk, "?", ".?")
Reg.Pattern = "^" & Isk
Reg.IgnoreCase = True

iL = LBound(DannIn, 1): iU = UBound(DannIn, 1)
ReDim Found(iL To iU) As Boolean  ' Явный тип Boolean

n = 0
For i = iL To iU
    ' Чтение значения из CondIn
    cellValue = CondIn(i, 1)
    
    ' Проверка на число и сравнение с 1
    If IsNumeric(cellValue) Then
        Found(i) = (CDbl(cellValue) = 1)
    Else
        Found(i) = False
    End If
    
    If Found(i) Then
        Found(i) = Reg.Test(DannIn(i, 1))  ' DannIn тоже одностолбцовый
        If Found(i) Then n = n + 1
    End If
Next
' конец Заменил блок с ошибкой
        If n = 0 Then
            i = MsgBox("Поиск по шаблону " & vbCr & vbCr & Isk0 & vbCr & vbCr & "неуспешен" & vbCr & vbCr & "Повторить ввод?", 33, "Поиск неуспешен")
            If i = 2 Then Exit Do
        Else
            ReDim Out(n)
            ii = -1
            For i = iL To iU
                If Found(i) Then
                    ii = ii + 1
                    Out(ii) = """" & Cbegin & (Rbegin - 1 + i) & """   " & DannIn(i, iL)
                End If
            Next
            
            If n = 1 Then
                Range(Replace(Split(Out(0), " ")(0), """", "")).Select
                Exit Do     ''''''''''''''''
            Else
                For i = 1 To n
                    Range(Replace(Split(Out(i - 1), " ")(0), """", "")).Activate
                    ii = MsgBox("Выбрать значение " & i & " из " & n & vbCr & vbCr & Out(i - 1), 35, "Найдено " & n & " совпадений " & """" & Isk0 & """")
                    If ii = 6 Then Exit Do  'For    ''''''''''''''''
                    If ii = 2 Then
                        Range("A1").Select
                        Exit For
                    End If
                Next
            End If
        End If
    Loop
End Sub



Sub ПоискВСтолбцеС_СПошаговымВыбором()
    Dim ws As Worksheet
    Dim searchTerm As String
    Dim lastRow As Long, i As Long
    Dim data As Variant
    Dim results As Collection
    Dim msg As String, userChoice As String
    Dim selectedIndex As Long, n As Long
    
    Dim Lfirst As Boolean: Lfirst = True  ' Флаг первой итерации


    ' Основной цикл — позволяет повторять поиск
    Do
        ' 1. Запрашиваем шаблон
        searchTerm = InputBox("Введите текст для поиска в столбце C:", "Поиск в столбце C", searchTerm)
        If searchTerm = "" Then
            i = MsgBox("Шаблон не введён." & vbCr & "Повторить ввод?", vbQuestion + vbYesNo, "Пустой ввод")
            If i = vbNo Then Exit Do  ' Выход из цикла
            GoTo NextIteration          ' Повтор ввода
        End If


        ' 2. На первой итерации — инициализация
        If Lfirst Then
            Set ws = ActiveSheet
            lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
            If lastRow < 1 Then
                MsgBox "В столбце C нет данных.", vbExclamation: Exit Do
            End If


            ' Считываем столбец C в массив
            If lastRow = 1 Then
                ReDim data(1 To 1, 1 To 1)
                data(1, 1) = ws.Cells(1, "C").Value
            Else
                data = ws.Range("C1:C" & lastRow).Value
            End If
            Lfirst = False
        End If

        ' 3. Собираем совпадения (номера строк)
        Set results = New Collection
        For i = 1 To UBound(data, 1)
            If Not IsEmpty(data(i, 1)) Then
                If InStr(1, LCase(CStr(data(i, 1))), LCase(searchTerm), vbTextCompare) > 0 Then
                    results.Add i
                End If
            End If
        Next i

        n = results.Count

        ' 4. Обрабатываем результаты
        If n = 0 Then
            i = MsgBox("Не найдено: """ & searchTerm & """" & vbCr & _
                      "Повторить ввод?", vbExclamation + vbYesNo, "Нет совпадений")
            If i = vbNo Then Exit Do
        Else
            ' Формируем сообщение со списком результатов (без номеров строк)
            msg = "Найдено " & n & " совпадений. Выберите номер (1–" & n & "), чтобы перейти к ячейке:" & vbCr & vbCr
            For i = 1 To n
                msg = msg & i & ". """ & data(results(i), 1) & """" & vbCr
            Next i

            ' Запрашиваем номер результата
            userChoice = InputBox(msg, "Выбор результата", 1)
            If userChoice = "" Then GoTo NextIteration  ' Отмена — новый цикл


            ' Проверяем ввод
            If Not IsNumeric(userChoice) Then
                MsgBox "Введите число!", vbExclamation: GoTo NextIteration
            End If
            selectedIndex = CLng(userChoice)
            If selectedIndex < 1 Or selectedIndex > n Then
                MsgBox "Номер вне диапазона (1–" & n & ")!", vbExclamation: GoTo NextIteration
            End If


            ' Переходим к выбранной ячейке
            ws.Cells(results(selectedIndex), "C").Select
            MsgBox "Переход выполнен.", vbInformation, "Ячейка выделена"
            Exit Do  ' Выходим после успешного перехода
        End If


NextIteration:
    Loop
End Sub


