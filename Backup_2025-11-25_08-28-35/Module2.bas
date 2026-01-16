Attribute VB_Name = "Module2"

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
        Isk0 = InputBox("Введите шаблон искомого слова ", , Isk0)
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
             Found(i) = True         '(CondIn(i, iL) = 1)
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

Global ClaimID As Range, ФИО As Range, Box As Range 'Сделал Клайм Global переменной
Public DeltaTime As Variant, Iskomoe As Variant

Sub Поиск_по_условию()
Application.WindowState = xlNormal
    Dim Cl As Range, Iskomoe$
    Iskomoe = InputBox("Введите несколько букв ФИО заёмщика", "Сообщение от Microsoft Excel", Iskomoe)
    If Iskomoe = "" Then Exit Sub
    Iskomoe = "*" & LCase(Iskomoe) & "*"
    For Each Cl In Range("C4:C" & Range("C4").End(xlDown).Row)
      If LCase(Cl) Like Iskomoe And Cl.Offset(, 48) = "1" Then Cl.Select
    Next
   
End Sub


'Крайняя действующая рабочая процедура (27.06.25)

Sub СТАРТ_3()  'Жму кнопу "СТАРТ" но ЛУЧШЕ клавишу перехвата {UP} - стрелка вверх
    Dim objNetwork: Set objNetwork = CreateObject("WScript.Network")
    Dim strComputerName: strComputerName = objNetwork.ComputerName  '// имя компа
    Dim username: username = objNetwork.username  '// имя пользователя
'    Dim Box '// № Коробки ФКБ
'     Box = 204     '                 <-<-<-------МЕНЯЮ № КОРОБКИ в "1м МЕСТЕ!
'     Range("B2") = Box

'    Box = Range("B2")             '         <-<-<-------МЕНЯЮ № КОРОБКИ в ЯЧЕЙКЕ. так проще.
Skip:

'CreateObject("WScript.Shell").Popup "Вставьте доки в сканер, окно закроется ч/з 1 сек.", 1, "Microsoft Excel", 48

    TimeStart = Timer   'переменная для тайминга процедуры
'    Sleep (400) 'Спим...
    SetCursorPos 20, 1023           'клик на СКАНЕР
        mouse_event &H2, 0, 0, 0, 0 'нажал ЛЕВУЮ кнопку мыши, H2
        mouse_event &H4, 0, 0, 0, 0 'отпустил ЛЕВУЮ кнопку мыши, H4
    Sleep (100)  'Спим...
        mouse_event &H2, 0, 0, 0, 0
        mouse_event &H4, 0, 0, 0, 0
        
'    Workbooks("Итог_ФКБ_Лазарев.xlsm").Sheets("Лист1").Activate ' Активирую книгу. НЕ АКТИВИРУЕТСЯ. (вар-2)
    AppActivate ("Итог_ФКБ_Лазарев.xlsm - Excel")  ' Активирую книгу. АКТИВИРУЕТСЯ. (вар-3)
    ActiveCell.Offset(0, -1).Activate 'Перехожу на одну ячейку левее активной и активируем ее.
    Set ClaimID = ActiveCell
    Set ФИО = ActiveCell.Offset(, 1)
    Set Box = Range("B2")
'1. Вывод статистики в окно Immediate:
'            Debug.Print "Пользователь: " & username & "  |" & "   Комп: " & strComputerName & "  |  " & "Задача #5: Сканирование ФКБ  |   Коробка №: " & Box
'            Debug.Print "Время старта обработки строки с Клаймом: " & Now
'            Debug.Print "Набранные символы в поле Окна ввода данных: " & Isk0 'Iskomoe
'            Debug.Print "Claim ID: " & ClaimID
'            Debug.Print "ФИО: " & ФИО
'            Debug.Print "Коробка №: " & Box
    
    Start = Timer 'Спим...
    Do While Timer < Start + 1 '0.5 = полсекунды   '  ЗАЧЕМ Спим ????????????????????????????????????????????????????????????
        DoEvents
    Loop
       
'2  АВТОЗАПОЛНЯЮ строку активной ячейки:
    If Not Intersect(ActiveCell, Range("B4:B77777")) Is Nothing Then
    Cells(ActiveCell.Row, 38).Value = " не учтено "
'    Cells(ActiveCell.Row, 37).Value = Now ' Именно Now, а не Date!
'    If Cells(ActiveCell.Row, 44).Value = 0 And IsEmpty(Cells(ActiveCell.Row, 44)) = True Then  'Было
        If Cells(ActiveCell.Row, 44).Value = 0 Then 'Стало
                Cells(ActiveCell.Row, 44).Value = Box 'ТУТА <--МЕНЯЮ № КОРОБКИ только в 1м МЕСТЕ!
            Else
                Cells(ActiveCell.Row, 44).Value = Cells(ActiveCell.Row, 44).Value & ", " & Box 'Ставит запятую
'               Cells(ActiveCell.Row, 44).Value = Cells(ActiveCell.Row, 44).Value & "," & Box 'НЕ ставит запятую
        End If
    Sheets("Лист1").Cells(ActiveCell.Row, 42).FormulaLocal = "=ЕСЛИ(AY4:AY77777=1;""Зарегистрирован"";""Отсутствует"")"
    Sheets("Лист1").Cells(ActiveCell.Row, 42) = Sheets("Лист1").Cells(ActiveCell.Row, 42).Value '  вставил значение формулы
    Cells(ActiveCell.Row, 45).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 45).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 46).Value = Now ' Именно Now, а не Date!   или как ниже
    Cells(ActiveCell.Row, 46).NumberFormat = "dd.mm.yyyy hh:mm:ss"   '   или как ниже
'        Worksheets("Лист1").Cells(ActiveCell.Row, 38).ClearContents 'Чищу ячейку
'    Sheets("Лист1").Cells(ActiveCell.Row, 46).FormulaLocal = "=ЕСЛИ(AS:AS=СЕГОДНЯ();AK:AK)" ' ниже вставлю значение
'    Sheets("Лист1").Cells(ActiveCell.Row, 46) = Sheets("Лист1").Cells(ActiveCell.Row, 46).Value '  вставил значение формулы
        Sheets("Лист1").Cells(ActiveCell.Row, 48).FormulaArray = _
        "=""Скорость: ""&ROUND(((COUNTIF(C45,TODAY()))/IF(ISNA(MATCH(TODAY(),R4C45:R77777C45,0)),""нет"",MAX(IF(R4C45:R77777C45=TODAY(),R4C46:R77777C46,))-MIN(IF((R4C45:R77777C45=TODAY())*(R4C46:R77777C46>0),R4C46:R77777C46,99^99))))/24,2)&"" строки в час"""
        Sheets("Лист1").Cells(ActiveCell.Row, 48) = Sheets("Лист1").Cells(ActiveCell.Row, 48).Value
    End If
    
'   2.2.ЗАКРАШИВАЮ ЯЧЕЙКИ в зависимости от полученного значения в рез-те применения формулы:
    Dim xCell As Range ''nj выражение объявляет переменную xCell с типом Range, что позволяет работать с ячейкой или диапазоном ячеек через эту переменную.
'   Объект Range может представлять одну ячейку, несколько ячеек (в том числе несмежные ячейки или наборы несмежных ячеек) или целый лист
    Dim CommentValue As String 'тип данных String, что означает, что переменная будет хранить строки. CommentValue-Значение комментария
    Dim CommentRange As Range ' означает объявление переменной CommentRange как диапазона
    Set CommentRange = Sheets("Лист1").Cells(ActiveCell.Row, 42) ' означает, что CommentRange будет представлять диапазон ячеек с адресом ActiveCell.Row, 42

    For Each xCell In CommentRange ' означает цикл для каждой ячейки в диапазоне CommentRange.
    CommentValue = xCell.Value 'означает, что переменная CommentValue получает значение ячейки xCel
    Select Case CommentValue 'Select Case проверяемое_выражение
        Case "Отсутствует" 'если проверяемое_выражение = Отсутствует
        xCell.Interior.Color = RGB(255, 0, 0) ' крашу в красный
        Case "Зарегистрирован" ' соответственно...
        xCell.Interior.Color = RGB(0, 255, 0) 'в зеленый
    End Select
    Next xCell
    
''   МЕНЯЮ ФОРМУЛЫ НА ЗНАЧЕНИЯ ----------
'    Dim smallrng As Range
'    ActiveCell.Offset(0, 35).Activate
'    ActiveCell.Resize(1, 11).Select 'выделяю 11 ячеек справа от активной ячейки вместе с активной и циклом
'        For Each smallrng In Selection.Areas 'преобразую формулы в значения в выделенном диапазоне
'            smallrng.Value = smallrng.Value
'        Next smallrng
    
    Cells(ActiveCell.Row, 2).Select 'перехожу во вторую ячейку строки из активной
    
    Range("AS3").Value = Format(Now, "HH:mm:ss") '<--- Вношу значение текущего времени в часы.
    
    Dim ВремяСканирования As String
    ВремяСканирования = Range("AZ3")
    
                Start = Timer ' Пауза для ................................
                               Do While Timer < Start + 0.6
                                   DoEvents
                               Loop
    
    
    
    f1_Таймер.Show
        DoEvents

            Application.Wait Now + TimeSerial(0, 0, 0.7)
                
'    Start = Timer
'    Do While Timer < Start + ВремяСканирования '<-------- ДАЮ ВРЕМЯ СКАНЕРУ ОТРАБОТАТЬ (13 листов это + 22 сек)
'          DoEvents
'    Loop ' Конец паузы.



'5  НАЧИНАЮ ПЕРЕИМЕНОВКУ ПАПКИ ----------------------------------------------------------------------------------------
    Dim objFSO As Object, objFolder As Object
    Dim sFolderName As String, sNewFolderName As String
    sFolderName = "C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\СКАНЫ_в работе\Новая папка\"  'имя исходной папки
    sNewFolderName = "Новая папка"  ' <- имя папки для переименования (только имя, без полного пути)
    'Создание объекта FileSystemObject:
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Проверка наличия папки по указанному пути:
    If objFSO.FolderExists(sFolderName) = False Then
      MsgBox "Нет такой папки", vbCritical, "ИНФОРМЕР ОТСУТСТВИЯ ПАПКИ"
      Exit Sub
    End If
    
'5.1    'ПЕРЕИМЕНОВКА ПАПКИ:
     Set objFolder = objFSO.GetFolder(sFolderName) 'Получение доступа к объекту Folder (папка)
    '  Назначение нового имени:
    objFolder.Name = ClaimID ' Обозвал папку значением ClaimID ' Тут бывает ошибка, когда прога вперед сканера идёт.
    
'5.2  Начал переименовку файла в папке ClaimID.............................................................................

    Dim oldFilePath As String
    Dim newFilePath As String
    Dim str As String
    Dim Str2 As String
    
    Dim a1, a2, a3, a4, a5

    a1 = "C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\СКАНЫ_в работе"
    a2 = ClaimID
    a3 = "Скан.pdf"
    a4 = ФИО
    a5 = ".pdf"
    
    str = a1 & "\" & a2 & "\" & a3
      Debug.Print "Имя файла Old: " & str
    oldFilePath = str
    Str2 = a1 & "\" & a2 & "\" & a2 & "_" & a4 & a5
      Debug.Print "Имя файла New: " & Str2
    newFilePath = Str2

'    On Error GoTo ErrorHandler ' Файл в папке еще не успел появиться. (например)
    Name oldFilePath As newFilePath  ' Это собственно сама переименовка файла
'   Debug.Print "Переименовка файла завершена!" '  MsgBox
''    Exit Sub
'ErrorHandler:
'    MsgBox "Ошибка переименовки файла: " & Err.Description

'''Конец переименовки файла в папке.....................................................................................
    
'7  Подключаю подпрограмму Проверки соотношения ПАПОК и ФАЙЛОВ
    Call Сколько_Папок_и_Файлов
'8  ПРОВЕРКА НА СОВПАДЕНИЕ ЗНАЧЕНИЙ ЯЧЕЕК. ЕСЛИ НЕ СОВПАЛИ- ОШИБКА.......................................................
    Dim СделалСегодня As Variant, проверка1 As Variant, проверка2 As Variant, проверка3 As Variant
    
    Range("AM2").FormulaR1C1 = "=COUNTIF(C45,TODAY())"   '=СЧЁТЕСЛИ($AS:$AS;СЕГОДНЯ())
    Range("AM2") = Range("AM2").Value
    СделалСегодня = Range("AM2").Value
    
'    Debug.Print " "
'    Debug.Print " ***************************************************** "
'    Debug.Print СделалСегодня  ' Значение Формулы в ячейке AM2: =СЧЁТЕСЛИ(AS:AS;СЕГОДНЯ()) FormulaR1C1 = "=COUNTIF(C45,TODAY())"
'    Debug.Print iCountFolders  ' Значение  iCountFolders
'    Debug.Print iCountFiles    ' Значение  iCountFiles
    
    Range("BB6").FormulaArray = "=IF(AND(COUNTIF(C45,TODAY())=R3C46:R3C47), ""да"",""нет"")" ' Формула в ячейке: {=ЕСЛИ(И(СЧЁТЕСЛИ($AS:$AS;СЕГОДНЯ())=AT3:AU3); "да";"нет")}
    проверка1 = Range("BB6").Value
    Range("BB6").ClearContents
    
    If iCountFolders = iCountFiles Then
        проверка2 = "да"
        Else: проверка2 = "нет"
    End If
    
    Range("BB8").FormulaArray = "=IF(AND(R2C39<R3C46:R3C47), ""да"",""нет"")" ' Формула в ячейке: {=ЕСЛИ(И(AM2<AT3:AU3); "да";"нет")}
    проверка3 = Range("BB8").Value
    Range("BB8").ClearContents
    
        If проверка1 = "да" Then
'            Debug.Print "Три значения равны/ошибок нет, продолжаем процедуру."
'            MsgBox "Зачения равны, ОК- продолжить процедуру", vbOKOnly, "Проверка трёх значений на равенство"
            Else
                If проверка2 = "нет" Then
                    Debug.Print "Три значения равны? - " & проверка1 & ".   " & "Кол-во сканов и папок совпадает? - "; проверка2
                    MsgBox "ВНИМАНИЕ, обнаружена ошибка!" & vbCr & vbNewLine & "Кол-во папок и сканов в них не совпадает." & vbCr & vbNewLine & "Программа будет прервана для внесения исправлений.", 48, "Чек совпадения количества строк, папок и файлов"
                    Exit Sub
                End If
                If проверка3 = "да" Then
                    Debug.Print "Три значения равны? - " & проверка1 & ".   " & "Кол-во сканов и папок совпадает? - " & проверка2 & ". Кол-во строк меньше кол-ва сканов и папок? - "; проверка3
                    MsgBox "ВНИМАНИЕ, обнаружена ошибка!" & vbCr & vbNewLine & "Кол-во строк меньше кол-ва сканов и папок, проверь автозаполнение текущей строки!" & vbCr & vbNewLine & "Программа будет прервана для внесения исправлений.", 48, "Чек совпадения количества строк, папок и файлов"
                    Exit Sub
                End If
        End If
'       ..................................................................................................................
    
'9      Вывод статистики в окно Immediate:
'    DeltaTime = Round(Timer - TimeStart, 2)
'    Dim DateEx As Date, DateEx2 As Date
'    DateEx = Range("AY3")
'    DateEx2 = Range("BB3")
'        Debug.Print "Данные для проверки: Папок = " & iCountFolders & ", Сканов = " & iCountFiles & ", Сделал за сегодня = " & Range("AM2").Value & " шт."
'        Debug.Print "Результат проверки: Три значения равны, ошибок нет."
'        Debug.Print "Запас строк/сканов: " & Range("AN2").Value & " шт."
'        Debug.Print "Время работы сканера: " & ВремяСканирования & " сек."
'        Debug.Print "Тайминг обработки Клайма: " & DeltaTime & " сек."
'        Debug.Print Range("AV3").Value  '"Скорость: "
'        Debug.Print "Оставшееся время до достижения 200 строк: " & Format(DateEx, "hh:nn:ss")
'        Debug.Print "200 строк будут сделаны в : " & Format(DateEx2, "hh:nn:ss")
'        Debug.Print "..........................................................."
'        Debug.Print " "
        
'    Const lSeconds As Long = 1  ' Пауза - 1 секунда
'        MessageBoxTimeOut 0, "Готово." & _
'           vbNewLine & vbCr & "Это окно закроется само" & _
'           vbNewLine & vbCr & "Клиент:  " & ClaimID & "  " & ФИО, "Сообщение от Microsoft Corporation: Процедура завершена", _
'        vbInformation + vbOKOnly, 0&, lSeconds * 1000
  
'  Dim WshShell As Object
'    Set WshShell = CreateObject("WScript.Shell")
'    WshShell.Popup "Все сделано" & vbCr & "Установленное в AZ3 время сканирования: " & ВремяСканирования & " сек.", 2, "Информационное сообщение от Microsoft Corporation", vbInformation    ' ждем 3 секунды
    
    
''10 Call Поиск_по_условию
'    Dim Cl As Range ', Iskomoe$
'        Iskomoe = InputBox("Введите несколько букв ФИО заёмщика", "Сообщение от Microsoft Excel", Iskomoe)
'        If Iskomoe = "" Then Exit Sub
'        Iskomoe = "*" & LCase(Iskomoe) & "*"
''            Debug.Print "Набранные символы в поле ввода данных: " & Iskomoe
'        For Each Cl In Range("C4:C" & Range("C4").End(xlDown).Row)
'          If LCase(Cl) Like Iskomoe And Cl.Offset(, 48) = "1" Then Cl.Select
'        Next
        
  

  Call Поиск_от_Angry_Old_Man
GoTo Skip

End Sub

Sub Сколько_Папок_и_Файлов()
    iCountFolders = 0     ' Обнулю значение переменной иначе будет удваиваться
    iCountFiles = 0       ' Обнулю значение переменной иначе будет удваиваться

    Dim iPath$ ', iCountFolders&, iCountFiles&
    Dim iFolder As Object, iFolderItem As Object
    iPath = "C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\СКАНЫ_в работе"
    Set iFolder = CreateObject("Shell.Application").Namespace(CVar(iPath))
    If Not (iFolder Is Nothing) Then
        Call NextFold(iFolder, iCountFiles, iCountFolders)
        Cells(3, 46).Value = iCountFolders
        Cells(3, 47).Value = iCountFiles
    Else
        MsgBox "такой папки нет", vbCritical, iPath
    End If
    If (Range("AT3").Value) <> (Range("AU3").Value) Then
    '        MsgBox "КОЛИЧЕСТВО ПАПОК И ФАЙЛОВ В КАТАЛОГЕ НЕ СОВПАДАЕТ!", vbExclamation
        End If
    '    Debug.Print "Папок: " & iCountFolders & ", Сканов: " & iCountFiles; " & Сделал строк за сегодня: " & Range("AM2").Value & " шт."
End Sub

Function NextFold(p As Variant, ByRef flCount As Long, ByRef fldCount As Long)
    If Not (p Is Nothing) Then
        For Each iFolderItem In p.Items
            If Not iFolderItem.IsFolder Then
                flCount = flCount + 1
            Else
                fldCount = fldCount + 1
                Call NextFold(CreateObject("Shell.Application").Namespace(CVar(iFolderItem.Path)), flCount, fldCount)
            End If
        Next iFolderItem
    End If
End Function


