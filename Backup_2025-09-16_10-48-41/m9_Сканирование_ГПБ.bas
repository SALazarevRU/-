Attribute VB_Name = "m9_Сканирование_ГПБ"
Option Explicit
Global iCountFolders22&, iCountFiles22&   'для считывания в процедуру Таймер
Global ClaimID_2 As Range, ФИО_2 As Range, Box_2 As Range   'Сделал Клайм Global переменной
Global Dosye_2 As Variant
Public DeltaTime_2 As Variant, Iskomoe_2 As Variant

Sub CallСканингГПБ()
    Call СканингГПБ(Nothing)
End Sub

Sub Проверка_кратности_5()
    If (cell.Value Mod 5) = 0 Then
        MsgBox "Число в ячейке кратно 5"
    End If
End Sub

Public Sub НомерКоробки(editBox As IRibbonControl, ByRef Text)
On Error GoTo Instruk
    Dim НомКоробки  As Long
    НомКоробки = ActiveSheet.Range("AP1")
    Text = "   " & НомКоробки
Instruk: Exit Sub
End Sub

Public Sub ВремяСканирования(editBox As IRibbonControl, Text As Variant)   ' При заполнении в боксе сразу меняется в ячейке!
    Dim editBox_ВремяСканирования As Variant
    editBox_ВремяСканирования = Text
    ActiveSheet.Range("AX1") = editBox_ВремяСканирования
'    Worksheets("Расширенный реестр").Range("AX1") = ВремяСкан
    Text = "   " & editBox_ВремяСканирования
End Sub

Public Sub CaptureText(editBox As IRibbonControl, Text As Variant) 'As Long 'для числового значения и As String для текстового
   Dim EditBoxТекст  As Variant 'для числового значения и As String для текстового
   EditBoxТекст = Text
   ActiveSheet.Range("AT1") = EditBoxТекст
End Sub



'обе четыре процедуры работаю в связке!
Sub СканингГПБ(control As IRibbonControl)  'Жму кнопу "Сканинг ГПБ" но ЛУЧШЕ клавишу перехвата {UP} - стрелка вверх
    
    If Worksheets("Расширенный реестр").Range("AX1").Value = "" Then
       Worksheets("Расширенный реестр").Range("AX1").Value = 8
    End If

    Dim result As Integer, result_2 As Integer
'    On Error GoTo Instruk

    ''Sub ArrPatt() ' Использую для Sub СканингГПБ() m9_Сканирование_ГПБ

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
            Found(i) = (CondIn(i, iL) = 1) ' ЭТО ДЛЯ ПОИСКА С УСЛОВИЕМ
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
''Sub ArrPatt() ' Конец


f1_Выбор_варианта_заполнения.Show ' ОТКРЫВАЮ ФОРМУ f1_Выбор_варианта_заполнения
        DoEvents
            Application.Wait Now + TimeSerial(0, 0, 0.7)
'Instruk2:
End Sub



Sub Показать_Форму_Выбор_варианта_заполнения(control As IRibbonControl)
     f1_Выбор_варианта_заполнения.Show
End Sub



Sub СканингГПБ_Часть2()
Dim TimeStart As Date
    
    TimeStart = Timer   'переменная для тайминга процедуры
'    Sleep (400) 'Спим...
    SetCursorPos 20, 1023           'клик на СКАНЕР
        mouse_event &H2, 0, 0, 0, 0 'нажал ЛЕВУЮ кнопку мыши, H2
        mouse_event &H4, 0, 0, 0, 0 'отпустил ЛЕВУЮ кнопку мыши, H4
    Sleep (100)  'Спим...
        mouse_event &H2, 0, 0, 0, 0
        mouse_event &H4, 0, 0, 0, 0
        
    AppActivate ("03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.xlsx - Excel")  ' Активирую книгу. АКТИВИРУЕТСЯ. (вар-3)
    
         Cells(1, 40).FormulaLocal = "=СЧЁТЕСЛИ(AO2:AO5689;СЕГОДНЯ())"    ' это Range("AN1") кол-во строк за сегодня
     Dosye_2 = Cells(1, 40).Value '    Сука, если в лэйбле не появляется значение глобальной переменной - СМОТРИ на ШИРИНУ ПОЛЯ лейбла !!!
    
            Application.Wait Now + TimeSerial(0, 0, 0.5)
             If (Range("AN1").Value Mod 1) = 0 Then ' Сохраняю книгу при условии.
    '        MsgBox "Число в ячейке кратно 5"
             ActiveWorkbook.Save
            
             f1_ФАЙЛ_СОХРАНЕН.Show 0
                Application.Wait Now + TimeValue("00:00:03")
                DoEvents
             Unload f1_ФАЙЛ_СОХРАНЕН
         End If
            
   '    f1_ТаймерГПБ.Show '<--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА  <--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА
    If MsgBox("По завершению сканирования нажмите ДА (Enter) для автопереименовки папки и скана в ней по имени клайма " & ClaimID_2 _
            & vbNewLine & "Спасибо!", vbYesNo, "© Рекомендация разработчика: ") = vbNo Then Exit Sub
        DoEvents
'            Application.Wait Now + TimeSerial(0, 0, 0.8)

'5  НАЧИНАЮ ПЕРЕИМЕНОВКУ ПАПКИ ----------------------------------------------------------------------------------------
    Dim objFSO As Object, objFolder As Object
    Dim sFolderName As String, sNewFolderName As String
    sFolderName = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе\Новая папка\"  'имя исходной папки
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
    objFolder.Name = ClaimID_2 ' Обозвал папку значением ClaimID ' Тут бывает ошибка, когда прога вперед сканера идёт.
    
'5.2  Начал переименовку файла в папке ClaimID.............................................................................

    Dim oldFilePath As String
    Dim newFilePath As String
    Dim str As String
    Dim Str2 As String
    
    Dim a1, a2, a3, a4, a5

    a1 = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе"
    a2 = ClaimID_2
    a3 = "Скан.pdf"
    
    a5 = ".pdf"
    
    str = a1 & "\" & a2 & "\" & a3
'      Debug.Print "Имя файла Old: " & str
    oldFilePath = str
    Str2 = a1 & "\" & a2 & "\" & a2 & a5
'      Debug.Print "Имя файла New: " & Str2
    newFilePath = Str2

    Name oldFilePath As newFilePath  ' Это собственно сама переименовка файла
    Set objFSO = Nothing
    
      Call Сколько_Папок_и_Файлов_в_ГПБ
    
     If Not gRibbon Is Nothing Then ' обновляю текстбоксы
        gRibbon.InvalidateControl "editBox_Строк"
         gRibbon.InvalidateControl "editBox_папок"
          gRibbon.InvalidateControl "editBox_файлов"
           gRibbon.InvalidateControl "editBox_Номер_коробки"
    End If
End Sub



Sub Сколько_Папок_и_Файлов_в_ГПБ()
'Dim iCountFolders22&, iCountFiles22&

iCountFolders22 = 0    ' Обнулю значение переменной иначе будет удваиваться
    iCountFiles22 = 0     ' Обнулю значение переменной иначе будет удваиваться

    Dim iPath$ ', iCountFolders&, iCountFiles&
    Dim iFolder As Object, iFolderItem As Object
    iPath = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе"   ' ПОД ЗАМЕНУ
    Set iFolder = CreateObject("Shell.Application").Namespace(CVar(iPath))
    If Not (iFolder Is Nothing) Then
        Call NextFold(iFolder, iCountFiles22, iCountFolders22)
      
        Workbooks("03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.xlsx").Worksheets("Расширенный реестр").Range("AR1").Value = iCountFiles22
'         MsgBox "КОЛИЧЕСТВО ПАПОК из ячейки AR1 " & Range("AR1").Value
         
         ActiveSheet.Range("AS1").Value = iCountFiles22
'         MsgBox "КОЛИЧЕСТВО СКАНОВ из ячейки AS1 " & Range("AS1").Value
    Else
        MsgBox "такой папки (C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе) - нет", vbCritical, iPath
    End If
'     '''ПРОВЕРКА СООТНОШЕНИЯ ПАПОК И СКАНОВ В ДИРЕКТОРИИ:
'    If iCountFolders22 + 1 <> (iCountFiles22) Then
'            MsgBox "КОЛИЧЕСТВО ПАПОК И ФАЙЛОВ В КАТАЛОГЕ НЕ СОВПАДАЕТ!", vbExclamation
'        End If
    '    Debug.Print "Папок: " & iCountFolders & ", Сканов: " & iCountFiles; " & Сделал строк за сегодня: " & Range("AM2").Value & " шт."
Set iFolder = Nothing
'Set iCountFiles22 = Nothing
'Set iCountFolders22 = Nothing
End Sub

Function NextFold(p As Variant, ByRef flCount As Long, ByRef fldCount As Long)
 Dim iFolderItem  As Variant
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
Sub пп()

End Sub

Public Sub строчек(editBox As IRibbonControl, ByRef Text)
'On Error GoTo Instruk
    Dim строк  As Long
    строк = ActiveSheet.Range("AN1")
    Text = "   " & строк
'Instruk:
'    Exit Sub
End Sub

Public Sub Папок(editBox As IRibbonControl, ByRef Text)
On Error GoTo Instruk
    Dim Папок  As Long
    Папок = ActiveSheet.Range("AR1")
    Text = "   " & Папок
Instruk:
    Exit Sub
End Sub

Public Sub Файлов(editBox As IRibbonControl, ByRef Text)
On Error GoTo Instruk
    Dim Файлов  As Long
    Файлов = ActiveSheet.Range("AS1")
    Text = "   " & Файлов
Instruk:
    Exit Sub
End Sub


Sub СканингГПБ_Часть3()             'эта часть для обработки строки в  "расчет выписка"
    Dim TimeStart As Date
        TimeStart = Timer           'переменная для тайминга процедуры
    SetCursorPos 20, 1023           'клик на СКАНЕР
        mouse_event &H2, 0, 0, 0, 0 'нажал ЛЕВУЮ кнопку мыши, H2
        mouse_event &H4, 0, 0, 0, 0 'отпустил ЛЕВУЮ кнопку мыши, H4
    Sleep (100)  'Спим...
        mouse_event &H2, 0, 0, 0, 0
        mouse_event &H4, 0, 0, 0, 0
        
    AppActivate ("03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.xlsx - Excel")  ' Активирую книгу. АКТИВИРУЕТСЯ. (вар-3)
    
        Cells(1, 40).FormulaLocal = "=СЧЁТЕСЛИ(AO2:AO5689;СЕГОДНЯ())"    ' это Range("AN1") кол-во строк за сегодня
        Dosye_2 = Cells(1, 40).Value ' Сука, если в лэйбле не появляется значение глобальной переменной - СМОТРИ на ШИРИНУ ПОЛЯ лейбла !!!
    
    Call Сколько_Папок_и_Файлов_в_ГПБ
            
    Application.Wait Now + TimeSerial(0, 0, 0.5)
            
         If (Range("AN1").Value Mod 1) = 0 Then ' Сохраняю книгу при условии.
    '        MsgBox "Число в ячейке кратно 5"
             ActiveWorkbook.Save
            
             f1_ФАЙЛ_СОХРАНЕН.Show 0
                Application.Wait Now + TimeValue("00:00:03")
                DoEvents
             Unload f1_ФАЙЛ_СОХРАНЕН
         End If
            
'   f1_ТаймерГПБ.Show '<--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА  <--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА
    Application.Wait Now + TimeValue("00:00:02")
    If MsgBox("По завершению сканирования нажмите ДА (Enter) для автопереименовки папки и скана в ней по имени клайма " & ClaimID_2 _
            & vbNewLine & "Спасибо!", vbYesNo, "© Рекомендация разработчика: ") = vbNo Then Exit Sub
        DoEvents
'            Application.Wait Now + TimeSerial(0, 0, 0.8)

'5  НАЧИНАЮ ПЕРЕИМЕНОВКУ ПАПКИ ----------------------------------------------------------------------------------------
    Dim objFSO As Object, objFolder As Object
    Dim sFolderName As String, sNewFolderName As String

    sFolderName = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе\Новая папка\"  'имя исходной папки
    sNewFolderName = "Новая папка"  ' <- имя папки для переименования (только имя, без полного пути)
    'Создание объекта FileSystemObject:
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Проверка наличия папки по указанному пути:
    If objFSO.FolderExists(sFolderName) = False Then
      MsgBox "Нет такой папки", vbCritical, "ИНФОРМЕР ОТСУТСТВИЯ ПАПКИ"
      Exit Sub
    End If
    
'5.1'ПЕРЕИМЕНОВКА ПАПКИ:
     Set objFolder = objFSO.GetFolder(sFolderName) 'Получение доступа к объекту Folder (папка)
    'Назначение нового имени:
     objFolder.Name = ClaimID_2 ' Обозвал папку значением ClaimID ' Тут бывает ошибка, когда прога вперед сканера идёт.
    
'5.2  Начал переименовку файла в папке ClaimID.............................................................................

    Dim oldFilePath As String
    Dim newFilePath As String
    Dim str As String
    Dim Str2 As String
    
    Dim a1, a2, a3, a4, a5, a7

    a1 = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе"
    a2 = ClaimID_2
    a7 = "расчет выписка"
    a3 = "Скан.pdf"
    a5 = ".pdf"
    
    str = a1 & "\" & a2 & "\" & a3
'      Debug.Print "Имя файла Old: " & str
    oldFilePath = str
    Str2 = a1 & "\" & a2 & "\" & a7 & a5
'      Debug.Print "Имя файла New: " & Str2
    newFilePath = Str2

'    On Error GoTo ErrorHandler ' Файл в папке еще не успел появиться. (например)
    Name oldFilePath As newFilePath  ' Это собственно сама переименовка файла
    Set objFSO = Nothing
         
    Call Сколько_Папок_и_Файлов_в_ГПБ
           
      If Not gRibbon Is Nothing Then ' обновляю текстбоксы
        gRibbon.InvalidateControl "editBox_Строк"
         gRibbon.InvalidateControl "editBox_папок"
          gRibbon.InvalidateControl "editBox_файлов"
           gRibbon.InvalidateControl "editBox_Номер_коробки"
    End If
End Sub


Sub СканингГПБ_Часть4() 'для переименовки скана в расчет выписка
Dim TimeStart As Date
    TimeStart = Timer   'переменная для тайминга процедуры
'    Sleep (400) 'Спим...
    SetCursorPos 20, 1023           'клик на СКАНЕР
        mouse_event &H2, 0, 0, 0, 0 'нажал ЛЕВУЮ кнопку мыши, H2
        mouse_event &H4, 0, 0, 0, 0 'отпустил ЛЕВУЮ кнопку мыши, H4
    Sleep (100)  'Спим...
        mouse_event &H2, 0, 0, 0, 0
        mouse_event &H4, 0, 0, 0, 0
        
    AppActivate ("03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.xlsx - Excel")  ' Активирую книгу. АКТИВИРУЕТСЯ. (вар-3)
    
         Cells(1, 40).FormulaLocal = "=СЧЁТЕСЛИ(AO2:AO5689;СЕГОДНЯ())"    ' это Range("AN1") кол-во строк за сегодня
     Dosye_2 = Cells(1, 40).Value '    Сука, если в лэйбле не появляется значение глобальной переменной - СМОТРИ на ШИРИНУ ПОЛЯ лейбла !!!
    
    
            Application.Wait Now + TimeSerial(0, 0, 0.5)
            
    If (Range("AN1").Value Mod 1) = 0 Then
'        MsgBox "Число в ячейке кратно 5"
         ActiveWorkbook.Save
'         CreateObject("WScript.Shell").Popup "Сохранил...", 1, "Сообщение о резервном копировании файла", 48
            f1_ФАЙЛ_СОХРАНЕН.Show 0
              Application.Wait Now + TimeValue("00:00:03")
              DoEvents
            Unload f1_ФАЙЛ_СОХРАНЕН
    End If
    
            
    '    f1_ТаймерГПБ.Show '<--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА  <--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА
    If MsgBox("Когда сканирование завершится - нажмите ДА (Enter) для автопереименовки папки и скана в ней по имени клайма " & ClaimID_2 _
            & vbNewLine & "Спасибо!", vbYesNo, "© Рекомендация разработчика: ") = vbNo Then Exit Sub
        DoEvents
'            Application.Wait Now + TimeSerial(0, 0, 0.8)

'5  НАЧИНАЮ ПЕРЕИМЕНОВКУ ПАПКИ ----------------------------------------------------------------------------------------
    Dim objFSO As Object, objFolder As Object
    Dim sFolderName As String, sNewFolderName As String
    sFolderName = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе\Новая папка\"  'имя исходной папки
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
    objFolder.Name = ClaimID_2 ' Обозвал папку значением ClaimID ' Тут бывает ошибка, когда прога вперед сканера идёт.
    
'5.2  Начал переименовку файла в папке ClaimID.............................................................................

    Dim oldFilePath As String
    Dim newFilePath As String
    Dim str As String
    Dim Str2 As String
    
    Dim a1, a2, a3, a4, a5, a7

    a1 = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе"
    a2 = ClaimID_2
    a7 = "расчет"
    a3 = "Скан.pdf"
    
    a5 = ".pdf"
    
    str = a1 & "\" & a2 & "\" & a3
'      Debug.Print "Имя файла Old: " & str
    oldFilePath = str
    Str2 = a1 & "\" & a2 & "\" & a7 & a5
'      Debug.Print "Имя файла New: " & Str2
    newFilePath = Str2

'    On Error GoTo ErrorHandler ' Файл в папке еще не успел появиться. (например)
    Name oldFilePath As newFilePath  ' Это собственно сама переименовка файла
    Set objFSO = Nothing
    
    Call Сколько_Папок_и_Файлов_в_ГПБ
    
   If Not gRibbon Is Nothing Then ' обновляю текстбоксы
        gRibbon.InvalidateControl "editBox_Строк"
         gRibbon.InvalidateControl "editBox_папок"
          gRibbon.InvalidateControl "editBox_файлов"
           gRibbon.InvalidateControl "editBox_Номер_коробки"
    End If
End Sub

Sub СканингГПБ_Часть5() 'для переименовки скана в выписка
Dim TimeStart As Date
    TimeStart = Timer   'переменная для тайминга процедуры
'    Sleep (400) 'Спим...
    SetCursorPos 20, 1023           'клик на СКАНЕР
        mouse_event &H2, 0, 0, 0, 0 'нажал ЛЕВУЮ кнопку мыши, H2
        mouse_event &H4, 0, 0, 0, 0 'отпустил ЛЕВУЮ кнопку мыши, H4
    Sleep (100)  'Спим...
        mouse_event &H2, 0, 0, 0, 0
        mouse_event &H4, 0, 0, 0, 0
        
    AppActivate ("03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.xlsx - Excel")  ' Активирую книгу. АКТИВИРУЕТСЯ. (вар-3)
    
         Cells(1, 40).FormulaLocal = "=СЧЁТЕСЛИ(AO2:AO5689;СЕГОДНЯ())"    ' это Range("AN1") кол-во строк за сегодня
     Dosye_2 = Cells(1, 40).Value '    Сука, если в лэйбле не появляется значение глобальной переменной - СМОТРИ на ШИРИНУ ПОЛЯ лейбла !!!
    
    
            Application.Wait Now + TimeSerial(0, 0, 0.5)
            
    If (Range("AN1").Value Mod 1) = 0 Then
'        MsgBox "Число в ячейке кратно 5"
         ActiveWorkbook.Save
'         CreateObject("WScript.Shell").Popup "Сохранил...", 1, "Сообщение о резервном копировании файла", 48
         f1_ФАЙЛ_СОХРАНЕН.Show 0
              Application.Wait Now + TimeValue("00:00:03")
              DoEvents
            Unload f1_ФАЙЛ_СОХРАНЕН
    End If

            
    f1_ТаймерГПБ.Show '<--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА  <--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА <--------  РАБОТА СКАНЕРА
    
        DoEvents
            Application.Wait Now + TimeSerial(0, 0, 0.8)

'5  НАЧИНАЮ ПЕРЕИМЕНОВКУ ПАПКИ ----------------------------------------------------------------------------------------
    Dim objFSO As Object, objFolder As Object
    Dim sFolderName As String, sNewFolderName As String
    sFolderName = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе\Новая папка\"  'имя исходной папки
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
    objFolder.Name = ClaimID_2 ' Обозвал папку значением ClaimID ' Тут бывает ошибка, когда прога вперед сканера идёт.
    
'5.2  Начал переименовку файла в папке ClaimID.............................................................................

    Dim oldFilePath As String
    Dim newFilePath As String
    Dim str As String
    Dim Str2 As String
    
    Dim a1, a2, a3, a4, a5, a7

    a1 = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\СКАНЫ_в_работе"
    a2 = ClaimID_2
    a7 = "выписка"
    a3 = "Скан.pdf"
    
    a5 = ".pdf"
    
    str = a1 & "\" & a2 & "\" & a3
'      Debug.Print "Имя файла Old: " & str
    oldFilePath = str
    Str2 = a1 & "\" & a2 & "\" & a7 & a5
'      Debug.Print "Имя файла New: " & Str2
    newFilePath = Str2

'    On Error GoTo ErrorHandler ' Файл в папке еще не успел появиться. (например)
    Name oldFilePath As newFilePath  ' Это собственно сама переименовка файла
    Set objFSO = Nothing
    
    Call Сколько_Папок_и_Файлов_в_ГПБ
    
      If Not gRibbon Is Nothing Then ' обновляю текстбоксы
        gRibbon.InvalidateControl "editBox_Строк"
         gRibbon.InvalidateControl "editBox_папок"
          gRibbon.InvalidateControl "editBox_файлов"
           gRibbon.InvalidateControl "editBox_Номер_коробки"
    End If
End Sub



