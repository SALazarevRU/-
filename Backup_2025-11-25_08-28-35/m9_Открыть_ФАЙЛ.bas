Attribute VB_Name = "m9_Открыть_ФАЙЛ"
Sub ОткрытьФайлОтчетаФабрика(control As IRibbonControl)
    Dim pathОтчета As String
    Dim pathОтчета_с_currentmonthname As String
    Dim currentmonthname As String
    Dim fl As String
    Dim wBook As Workbook
    
'    Application.AskToUpdateLinks = False   'Чтобы отключить сообщение «Нам не удалось обновить связи» НЕ РАБОТАЕТ
    Application.DisplayAlerts = False      'Чтобы отключить сообщение «Нам не удалось обновить связи»
'    currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
    currentmonthname = Format(Date, "mmmm") ' получаем полное название текущего месяца в Мммм
    
'    MsgBox "Имя текущего месяца: " & currentmonthname, vbInformation, "F1. Current month name"
    
    pathОтчета = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\Для руководителей\Операционная отчетность фабрики\2025\Отчет по фабрикe.New (Июль 2025).xlsx" ' полное имя Отчета по клаймам
   
    pathОтчета_с_currentmonthname = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\Для руководителей\Операционная отчетность фабрики\2025\" & "Отчет по фабрикe.New (" & currentmonthname & " 2025).xlsx"
        
'  проверка, открыт ли наш файл (без специальной функции Function IsBookOpen(wbFullName As String) As Boolean)
    Dim sWBName As String
    Dim bChk As Boolean
    
    sWBName = pathОтчета_с_currentmonthname
    
    For Each wBook In Workbooks
        If wBook.Name = sWBName Then
            Set wBook = Workbooks("Отчет по фабрикe.New (" & currentmonthname & " 2025).xlsx")
            wBook.Windows(1).Activate 'вытягиваем на первый план
        End If
    Next
    
    If bChk = False Then
        Set wBook = Workbooks.Open(FileName:=pathОтчета_с_currentmonthname)  ' открыл
    End If
 '  конец кода проверки, открыт ли наш файл
    
    With Application  ' размещаю окно  в координатах:
        .WindowState = xlNormal
        .Width = 1420 ' ШИРИНА окна
        .Height = 507 ' ВЫСОТА окна
        .Left = 0     ' ЛЕВАЯ точка
        .Top = 0      ' ВЕРХНЯЯ точка
    End With
    
'   Application.AskToUpdateLinks = True 'Чтобы включить сообщение «Нам не удалось обновить связи»
    Application.DisplayAlerts = True
    Range("B1").End(xlDown).Select 'прокрутка до последней заполненной ячейки столбца Б
End Sub

Public Sub ОткрытьФайлДинамика2025(control As IRibbonControl)

    Workbooks.Open FileName:="Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx"
    
    Application.ScreenUpdating = False        'оператор для отключения обновления экрана (=ускорение), он же фоновый режим(?))
   
        With Application  ' размещаю окно  в координатах:
            .WindowState = xlNormal
            .Width = 1420 ' ШИРИНА окна
            .Height = 807 ' ВЫСОТА окна
            .Left = 0     ' ЛЕВАЯ точка
            .Top = 0      ' ВЕРХНЯЯ точка
        End With
    
    Application.ScreenUpdating = True        'включаю обновление экрана.

End Sub

Public Sub Открыть_файл_ФКБ_МОЙ()
Workbooks.Open FileName:="C:\Users\s.lazarev\Desktop\2_Быстроденьги_сканинг\Итог_ФКБ 1 2 3 элек+ бумаж_МОЙ_NEW.xlsx"

  Application.ScreenUpdating = False        'оператор VBA для отключения обновления экрана (=ускорение), он же фоновый режим(?))
   
        With Application  ' размещаю окно  в координатах:
            .WindowState = xlNormal
            .Width = 1420 ' ШИРИНА окна
            .Height = 300 ' ВЫСОТА окна
            .Left = 0     ' ЛЕВАЯ точка
            .Top = 230      ' ВЕРХНЯЯ точка
        End With
    
    Application.ScreenUpdating = True        'включаю обновление экрана.
    
'   снимаю фильтры на активном листе [Итог.xlsx]:
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
        End If
    ActiveSheet.Cells.Rows.Hidden = False
    ActiveWindow.ScrollRow = 2

End Sub

Sub ОткрытьОтчетПОклаймам(control As IRibbonControl)  '  ПЕРЕДЕЛАТЬ НА ТЕКУЩ МЕСЯЦ
'    задаю переменные:
        Dim GCell As Range
        Dim txt$
        Dim wBook As Workbook
        Dim fl As Boolean
    
    If fl Then ' если уже был открыт мною
        Set wBook = Workbooks("Отчет по клаймам за сентябрь 2025.xlsx")
        wBook.Windows(1).Activate 'вытягиваем на первый план
    Else
        Set wBook = Workbooks.Open(FileName:="Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\Отчет по клаймам за ноябрь 2025.xlsx")  ' открыл
    End If
    
    With Application  ' размещаю окно  в координатах:
        .WindowState = xlNormal
        .Width = 1420 ' ШИРИНА окна
        .Height = 507 ' ВЫСОТА окна
        .Left = 0     ' ЛЕВАЯ точка
        .Top = 0      ' ВЕРХНЯЯ точка
    End With
'   снимаю фильтры на активном листе [Итог.xlsx]:
    If wBook.ActiveSheet.FilterMode = True Then
         wBook.ActiveSheet.ShowAllData
    End If
End Sub

' Открыть MyБлокнот - Файл Записки.Txt========================================================================================
'работает в связке с апи функциями PtrSafe Function FindWindow и PtrSafe Function SetWindowPos

Public Sub ОткрытьMyБлокнот(control As IRibbonControl)
'Sub ОткрытьMyБлокнот()
    Dim FilePath As String
    Dim hwnd As LongPtr
   
    ' Укажите путь к вашему текстовому файлу
    FilePath = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\Записная Книжка.txt" ' Замените на реальный путь
 Application.ScreenUpdating = False
    ' Открываем блокнот с файлом
    shell "notepad.exe " & FilePath, vbNormalFocus
    
    ' Ждем, пока блокнот откроется
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Находим окно блокнота
    hwnd = FindWindow("Notepad", vbNullString)

    ' Устанавливаем размеры окна блокнота
    If hwnd <> 0 Then
        SetWindowPos hwnd, 0, 1100, 300, 500, 700, SWP_NOZORDER Or SWP_NOACTIVATE
    Else
        MsgBox "Не удалось найти окно Блокнота."
    End If
    
    Application.ScreenUpdating = True
End Sub

' Открыть MyБлокно End========================================================================================

Public Sub ОткрытьАвторизация(control As IRibbonControl)
'Sub ОткрытьАвторизация()
   Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev"
       username = Environ("UserName")  ' Получаем имя пользователя.
        If username = SpecifiedUserName Then
            Application.ScreenUpdating = False
            ' Укажите путь к вашему текстовому файлу
              Workbooks.Open FileName:="C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\Авторизация.xlsx" ' Замените на реальный путь
            Application.ScreenUpdating = True
       Else
                MsgBox "Sorry, ресурс заблокирован" & vbNewLine & "Недостаточно прав доступа  ", 48, "Информация о блокировке доступа к выполнению программы"
            Exit Sub
        End If
End Sub
