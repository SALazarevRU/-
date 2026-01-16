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

Sub Поиск_от_Angry_Old_Man_БЕЗУСЛОВНЫЙ() ' КОСТЫЛЬ В 37 СТРОКЕ
    Dim title
     title = "Фамилия Имя Отчество должника"
    Dim BegDann As Range: Set BegDann = Range("B2")
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

