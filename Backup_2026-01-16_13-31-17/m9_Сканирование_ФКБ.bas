Attribute VB_Name = "m9_Сканирование_ФКБ"

Global ClaimID As Range, ФИО As Range, Box As Range 'Сделал Клайм Global переменной видимой отовсюду в проекте
Public DeltaTime As Variant, Iskomoe As Variant
Public CancelScan As Boolean

Public Isk0 As Variant

'Sub Поиск_от_Angry_Old_Man_ФКБ()
Sub Поиск_от_Angry_Old_Man_ФКБ(control As IRibbonControl)  'ВЕДЕТ СЕБЯ ТАК КАК НУЖНО - ПРИ ОТМЕНЕ ОН НЕ ЗАПУСКАЕТ БЕСКОНЕЧНЫЙ ЦИКЛ
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



'Крайняя действующая рабочая процедура (27.06.25)

Sub СТАРТ_СКАНИНГА_ФКБ()  'Жму кнопу "СТАРТ" но ЛУЧШЕ клавишу перехвата {UP} - стрелка вверх

    TimeStart = Timer   'переменная для тайминга процедуры
'    Sleep (400) 'Спим...
    SetCursorPos 20, 1023           'клик на СКАНЕР
        mouse_event &H2, 0, 0, 0, 0 'нажал ЛЕВУЮ кнопку мыши, H2
        mouse_event &H4, 0, 0, 0, 0 'отпустил ЛЕВУЮ кнопку мыши, H4
    Sleep (100)  'Спим...
        mouse_event &H2, 0, 0, 0, 0
        mouse_event &H4, 0, 0, 0, 0
        
    AppActivate ("Итог_ФКБ_Лазарев.xlsm - Excel")  ' Активирую книгу. АКТИВИРУЕТСЯ. (вар-3)
    ActiveCell.Offset(0, -1).Activate 'Перехожу на одну ячейку левее активной и активируем ее.
    Set ClaimID = ActiveCell
    Set ФИО = ActiveCell.Offset(, 1)
    Set Box = Range("B2")

    Start = Timer 'Спим...
    Do While Timer < Start + 1 '0.5 = полсекунды   '  ЗАЧЕМ Спим ????????????????????????????????????????????????????????????
        DoEvents
    Loop
       
'2  АВТОЗАПОЛНЯЮ строку активной ячейки:
    If Not Intersect(ActiveCell, Range("B4:B77777")) Is Nothing Then
    Cells(ActiveCell.Row, 38).Value = " не учтено "

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
        
      ' МОЯ РАБОЧАЯ ФОРМУЛА С ОШИБКОЙ ПРИ ПЕРВОЙ ИТЕРАЦИИ:
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
    
    
    Cells(ActiveCell.Row, 2).Select 'перехожу во вторую ячейку строки из активной
    
    Range("AS3").Value = Format(Now, "HH:mm:ss") '<--- Вношу значение текущего времени в часы.
    
    Dim ВремяСканирования As String
    ВремяСканирования = Range("AZ3")
    
                Start = Timer ' Пауза для ................................
                               Do While Timer < Start + 0.6
                                   DoEvents
                               Loop
    f1_ТаймерФКБ.Show
        DoEvents
            Application.Wait Now + TimeSerial(0, 0, 0.7)

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

    Name oldFilePath As newFilePath  ' Это собственно сама переименовка файла
    
'7  Подключаю подпрограмму Проверки соотношения ПАПОК и ФАЙЛОВ
    Call Сколько_Папок_и_Файлов
'8  ПРОВЕРКА НА СОВПАДЕНИЕ ЗНАЧЕНИЙ ЯЧЕЕК. ЕСЛИ НЕ СОВПАЛИ- ОШИБКА.......................................................
    Dim СделалСегодня As Variant, проверка1 As Variant, проверка2 As Variant, проверка3 As Variant
    
    Range("AM2").FormulaR1C1 = "=COUNTIF(C45,TODAY())"   '=СЧЁТЕСЛИ($AS:$AS;СЕГОДНЯ())
    Range("AM2") = Range("AM2").Value
    СделалСегодня = Range("AM2").Value

    
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
    
gRibbon.Invalidate ' Обновляет всю ленту
    DoEvents   ' Ключевая строка: даём Excel обработать обновления
'    Application.Wait Now + TimeValue("0:00:00") + 0.5 / 86400 ' Пауза (500 мс), чтобы не перегружать процессор
    Application.Wait Now + TimeValue("0:00:01") ' Пауза (1 сек), чтобы не перегружать процессор
    
  Call Поиск_от_Angry_Old_Man_БЕЗУСЛОВНЫЙ
'  Call Поиск_от_Angry_Old_Man_ФКБ(Nothing)
  Application.OnTime Now + TimeValue("00:00:01"), "СТАРТ_СКАНИНГА_ФКБ"
  
'GoTo Skip

End Sub


Sub ОСТАНОВИТЬ_СКАНИНГ_БЕЗОПАСНО()
    Dim nextTime As Date
    nextTime = Now
    
    On Error Resume Next
    Do While True
        Application.OnTime nextTime, "СТАРТ_СКАНИНГА_ФКБ", Schedule:=False
        If Err.Number <> 0 Then Exit Do
        nextTime = nextTime + TimeValue("00:00:01")
    Loop
    On Error GoTo 0
    
    MsgBox "Все запланированные запуски отменены.", vbInformation
End Sub

Public Sub Сколько_Папок_и_Файлов()
    Dim iCountFolders As Long  ' Явно указываем тип Long
    Dim iCountFiles As Long   ' Явно указываем тип Long
    Dim iPath As String
    Dim iFolder As Object, iFolderItem As Object
    
    ' Обнуляем счётчики
    iCountFolders = 0
    iCountFiles = 0
    
    ' Путь к папке
    iPath = "C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\СКАНЫ_в работе"
    
    ' Получаем объект папки через Shell
    Set iFolder = CreateObject("Shell.Application").Namespace(CVar(iPath))
    
    If Not (iFolder Is Nothing) Then
        Call NextFold(iFolder, iCountFiles, iCountFolders)
        
        ' Выводим результаты в ячейки
        Cells(3, 46).Value = iCountFolders  ' Папки > AP3
        Cells(3, 47).Value = iCountFiles      ' Файлы > AQ3
        
        ' Проверка несоответствия (если нужно)
        If (Range("AT3").Value <> Range("AU3").Value) Then
'            MsgBox "Количество папок и файлов не совпадает!", vbExclamation, "Внимание"
        End If
    Else
'        MsgBox "Папка не найдена: " & iPath, vbCritical, "Ошибка"
    End If
End Sub




' Рекурсивная функция для обхода папок
Private Sub NextFold(ByVal folder As Object, ByRef fileCount As Long, ByRef folderCount As Long)
    Dim item As Object
    
    If folder Is Nothing Then Exit Sub
    
    For Each item In folder.Items
        If Not item.IsFolder Then
            fileCount = fileCount + 1  ' Считаем файл
        Else
            folderCount = folderCount + 1  ' Считаем папку
            
            ' Рекурсивно обрабатываем вложенную папку
            Dim subfolder As Object
            Set subfolder = CreateObject("Shell.Application").Namespace(item.Path)
            If Not (subfolder Is Nothing) Then
                Call NextFold(subfolder, fileCount, folderCount)
            End If
        End If
    Next item
End Sub





Sub ПодсчетПодпапкок_СКЛАД_ФКБ()
    Dim fso As Object
    Dim folder As Object
    Dim rootPath As String
    Dim subfolderCount As Long
    Dim ws As Worksheet
    
    ' Укажите путь к вашей папке
    rootPath = "C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\СКАНЫ_СКЛАД  не сливал"
    
    ' Получаем лист "ппонФКБ" из активной книги
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("ппонФКБ")
    If ws Is Nothing Then
        MsgBox "Лист 'ппонФКБ' не найден в активной книге!", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Создаём объект FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Проверяем, существует ли папка
    If Not fso.FolderExists(rootPath) Then
        ws.Range("D56").Value = "Папка не найдена"
        MsgBox "Папка не найдена: " & rootPath, vbExclamation
        Exit Sub
    End If
    
    ' Получаем объект папки
    Set folder = fso.GetFolder(rootPath)
    
    ' Считаем подпапки (рекурсивно)
    subfolderCount = GetSubfolderCount(folder)
    
    ' Выводим результат в ячейку D56 листа "ппонФКБ"
    ws.Range("D56").Value = subfolderCount
    
'    MsgBox "Результат записан в ячейку D56 листа 'ппонФКБ'.", vbInformation
End Sub

' Рекурсивная функция для подсчёта подпапок
Function GetSubfolderCount(ByRef folder As Object) As Long
    Dim subfolder As Object
    Dim Count As Long
    
    Count = 0
    For Each subfolder In folder.Subfolders
        Count = Count + 1
        Count = Count + GetSubfolderCount(subfolder)  ' Рекурсия
    Next subfolder
    
    GetSubfolderCount = Count
End Function

