Attribute VB_Name = "m9_Динамические_строки"
Option Explicit

'Цель:создать полное имя активной книги(.xlsx) с именем пользователя
'Цель:создать полное имя этой книги (книги с макросом)с именем пользователя


Sub m9_Динамические_строки()
 Dim sFolderPath As String
    Dim sFileName As String
    Dim sFileExt As String
    Dim sNewFileName As String
    
   
'    sFolderPath = "C:\Users\Хозяин\Desktop\Бэкапы" & "\"
    sFolderPath = "C:\Users\s.lazarev\Desktop\Бэкапы" & "\"
    
    sFileName = ActiveWorkbook.Name
'    MsgBox sFileName, vbInformation
    sFileExt = VBA.Mid(sFileName, VBA.InStrRev(sFileName, ".", , vbTextCompare))
    sNewFileName = VBA.Replace(sFileName, sFileExt, "", , , vbTextCompare)
    
    sNewFileName = sFolderPath & sNewFileName & " (Backup) " & sDateTimeStamp & sFileExt
    
        Application.DisplayAlerts = False
    ActiveWorkbook.SaveCopyAs sNewFileName
'    MsgBox "Бэкап создан", vbInformation
        Application.DisplayAlerts = True
End Sub

Sub ggg() '   код возвращает имя активной книги, включая расширение файла
    Dim strWBName As String
    strWBName = ActiveWorkbook.Name
    MsgBox "имя активной книги, включая расширение файла: " & strWBName, vbInformation
End Sub

Sub ggg222() '  получить только имя без расширения, можно использовать функции Left и InStr:
    Dim strWBNameБезРасш  As String
    strWBNameБезРасш = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
    MsgBox "имя активной книги, без расширения файла: " & strWBNameБезРасш, vbInformation
End Sub

Sub GetActiveWorkbookPath() ' получить ПОЛНЫЙ ПУТЬ к активной книге используя свойство FullName объекта ActiveWorkbook. Это свойство возвращает путь к файлу, включая его название и расширение
    Dim activeWorkbookPath As String
    activeWorkbookPath = ActiveWorkbook.FullName
    MsgBox "Путь активной рабочей книги: " & activeWorkbookPath, vbInformation, "F1. ActiveWorkbook.Path"
End Sub

Sub Динамич_путь_с_user_name() 'код, меняющий путь к папке, вставляя переменную с именем пользователя:
    Dim pathОтчета As String
    Dim pathОтчета_с_ЮзерНэйм As String
    Dim user_name As String ' Объявляем переменную для имени пользователя
    
    user_name = Environ("UserName")  ' Получаем имя пользователя.
        
    pathОтчета = "C:\Users\Юзер\Desktop\Отчет по клаймам за май 2025.xlsx"
       
    pathОтчета_с_ЮзерНэйм = "C:\Users\" & user_name & "\Desktop\Отчет по клаймам за май 2025.xlsx" ' Заменяет каталог на каталог с user_name
   
    MsgBox "Новый Путь книги: " & pathОтчета_с_ЮзерНэйм, vbInformation, "F1. Динамический Path"

End Sub

Sub ШАБЛОН_1() 'полное название текущего месяца в нижнем регистре

    Dim currentmonthname As String
    currentmonthname = LCase(Format(Date, "mmmm")) ' возвращает полное название текущего месяца в нижнем регистре
      MsgBox "Полное название текущего месяца: " & currentmonthname, vbInformation, "F1.Название текущего месяца"
      
End Sub

Sub ШАБЛОН_2()  'Цель:получить полное имя Отчета по клаймам с динамическим текущим месяцем, проверить его на открытие и открыть если нет
    Dim pathОтчета As String
    Dim pathОтчета_с_currentmonthname As String
    Dim currentmonthname As String
    Dim fl As String
    Dim wBook As Workbook
     Dim времяинойработы As String
    Dim Задача As String
     Dim user_name As String
     
     user_name = Environ("UserName")

    currentmonthname = LCase(Format(Date, "mmmm")) ' получаем полное название текущего месяца в нижнем регистре
    pathОтчета = "C:\Users\Хозяин\Desktop\Отчет по клаймам за май 2025.xlsx" ' полное имя Отчета по клаймам с любым месяцем
'   pathОтчета = IsBookOpen("Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\Отчет по клаймам за июнь 2025.xlsx")' полное имя Отчета по клаймам
   
    pathОтчета_с_currentmonthname = "C:\Users\Хозяин\Desktop\" & "Отчет по клаймам за " & currentmonthname & " 2025.xlsx" ' новое имя Отчета по клаймам с текущим месяцем
   
    MsgBox "Новый Путь книги: " & vbNewLine & pathОтчета_с_currentmonthname, vbInformation, "F1. Динамический Path с текущим месяцем"
    
'      pathОтчета_с_currentmonthname = IsBookOpen("Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\ОТЧЕТ по клаймам\Отчет по клаймам за июнь 2025.xlsx")
        
'  проверка, открыт ли наш файл (без специальной функции Function IsBookOpen(wbFullName As String) As Boolean)
'  Sub checkWB() ' РАБОЧАЯ процедура для проверки, открыт ли файл в Excel, если да - на передний план, если нет - открыть и на передний план.
 
    Dim wb As Workbook
    Dim myWB As String
    Dim FileName As String
    
    FileName = pathОтчета_с_currentmonthname
    myWB = ("Отчет по клаймам за " & currentmonthname & " 2025.xlsx")
     
    For Each wb In Workbooks
        If wb.Name = myWB Then
            wb.Activate

            GoTo Punkt1
        End If
    Next wb
     Workbooks.Open FileName
 '  конец кода проверки, открыт ли наш файл
Punkt1:
 '  очистка переменных при необходимости  некоторые объекты и массивы
 'заполнить Лист "иное время" если нужно
     времяинойработы = InputBox("Если нужно,введите время иной работы", "Запрос данных", " ")
  
    If времяинойработы = "" Then
    GoTo Instruk
    
        Else 'заполняем иное время
           Worksheets("иное время").Select
            '   снимаю фильтры на активном листе:
                If ActiveSheet.FilterMode = True Then
                     ActiveSheet.ShowAllData
                End If
    
            Dim iLastRow As Long
                iLastRow = Cells(Rows.Count, 1).End(xlUp).Row
                Cells(iLastRow + 1, 1).Select
            
            With Selection.Validation '   в выбранной ячейке создать выпадающий список и выбрать нужное значение для неё
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Ващенко Оксана Васильевна,Лазарев Сергей Александрович"
                
            End With
            
        Select Case Environ("UserName")
        Case "v.petrov"
              Cells(iLastRow + 1, 1) = "Вася Петров"
        Case "s.lazarev"
            Cells(iLastRow + 1, 1) = "Лазарев Сергей Александрович"
        Case "g.sidorov"
              Cells(iLastRow + 1, 1) = "Гоша Сидоров"
        End Select
        
'                iLastRow = Cells(Rows.Count, 2).End(xlUp).Row
                 Cells(iLastRow + 1, 2).Value = "Иная работа" ' — ссылка на ячейку, расположенная в строке, следующей за последней, и во 2-м столбце.
                 
'                 iLastRow = Cells(Rows.Count, 3).End(xlUp).Row
                 Cells(iLastRow + 1, 3).Value = времяинойработы
                
'            If MsgBox("Заполнить дальше?", vbYesNo) = vbNo Then Exit Sub
'                iLastRow = Cells(Rows.Count, 4).End(xlUp).Row
                Cells(iLastRow + 1, 4).Value = " задача 34345"
                 

'                iLastRow = Cells(Rows.Count, 5).End(xlUp).Row
                Cells(iLastRow + 1, 5).Value = Date ' СЕГОДНЯ
    End If
    
Instruk:
End Sub
 

Sub ШАБЛОН_3()
   
End Sub

'  очистка переменных
'Integer,Long,Byte,Double,Decimal(уже не используется),Currency,Syngle,Date = по умолчанию имеют значение 0. Его и надо присваивать.
'String - =""
'Array - Erase Arr
'Object - Set obj = Nothing
'
'Variant в зависимости от назначенного типа. Но можно просто =0.

