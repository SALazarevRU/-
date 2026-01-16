Attribute VB_Name = "m9_Поиск_файлов_по_маске"
'Option Explicit
Option Compare Text 'Важно: если требуется, чтобы поиск не зависел от регистра символов в маске файла нужно поставить первой строкой в модуле директиву Option Compare Text

Sub Поиск_файлов_по_маске()  'РАБОТАЕТ
    Dim strDirPath, strMaskSearch, strFileName As String
    strDirPath = "Q:\LP2\Результаты сверки портфелей с августа 2020\" 'Папка поиска
'    strDirPath = "C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025\"
'    strMaskSearch = "*.xls*" 'Маска поиска
    strMaskSearch = "*Расширенный реестр*" 'Маска поиска
    'Получаем первый файл, соответствующий шаблону
    strFileName = Dir(strDirPath & strMaskSearch)
    'Перебираем файлы, пока они не закончатся
    Do While strFileName <> ""
'    MsgBox strFileName
    Debug.Print strFileName
    strFileName = Dir 'Следующий файл
    Loop
End Sub
'Важно: если требуется, чтобы поиск не зависел от регистра символов в маске файла (например, обнаруживались не только файлы .txt, но и .TXT и .Txt),
'нужно поставить первой строкой в модуле директиву Option Compare Text


Sub Поиск_файлов_по_маске_Регуляркой()
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objRegExp As Object
    
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = ".*xlsx"
    objRegExp.IgnoreCase = True
    
    Dim colFiles As Collection
    Set colFiles = New Collection
    RecursiveFileSearch "C:\Path\To\Your\Directory", objRegExp, colFiles, objFSO
    For Each f In colFiles
    Debug.Print (f)
    'Вставить код для работы с найденными файлами
    Next
    'Garbage Collection
    Set objFSO = Nothing
    Set objRegExp = Nothing
End Sub


Sub ЗагрузкаСпискаФайлов()

    ' Ищем файлы в заданной папке по заданной маске, и выводим на лист список их параметров. Просматриваются папки с заданной глубиной вложения.

    Dim coll As Collection, ПутьКПапке$, МаскаПоиска$, ГлубинаПоиска%
    ПутьКПапке$ = [r1]    ' берём из ячейки c1
    МаскаПоиска$ = [r2]    ' берём из ячейки c2
    ГлубинаПоиска% = Val([r3])    ' берём из ячейки c3
    If ГлубинаПоиска% = 0 Then ГлубинаПоиска% = 999    ' без ограничения по глубине

    ' считываем в колекцию coll нужные имена файлов

    Set coll = FilenamesCollection(ПутьКПапке$, МаскаПоиска$, ГлубинаПоиска%)

'    Application.ScreenUpdating = False    ' отключаем обновление экрана

    ' выводим результаты (список файлов, и их характеристик) на лист

    For i = 1 To coll.count    ' перебираем все элементы коллекции, содержащей пути к файлам

        НомерФайла = i
        ПутьКФайлу = coll(i)
        ИмяФайла = Dir(ПутьКФайлу)
        ДатаСоздания = FileDateTime(ПутьКФайлу)
        РазмерФайла = FileLen(ПутьКФайлу)

        ' выводим на лист очередную строку

        Range("a" & Rows.count).End(xlUp).Offset(1).Resize(, 5).Value = Array(НомерФайла, ИмяФайла, ПутьКФайлу, ДатаСоздания, РазмерФайла)

        ' если нужна гиперссылка на файл во втором столбце

        ActiveSheet.Hyperlinks.Add Range("b" & Rows.count).End(xlUp), ПутьКФайлу, "", "Открыть файл" & vbNewLine & ИмяФайла

        DoEvents    ' временно передаём управление ОС

    Next

End Sub

'PS: Найти подходящие имена файлов в коллекции можно при помощи следующей функции:

Function CollectionAutofilter(ByRef coll As Collection, ByVal filter$) As Collection

    ' Функция перебирает все элементы коллекции coll,

    ' оставляя лишь те, которые соответствуют маске filter$ (например, filter$="*некий текст*")

    ' Возвращает коллекцию, содержащую только подходящие элементы

    ' Если элементы не найдены - возвращается пустая коллекция (содержащая 0 элементов)

    On Error Resume Next: Set CollectionAutofilter = New Collection

    For Each item In coll

        If item Like filter$ Then CollectionAutofilter.Add item

    Next

End Function

Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                             Optional ByVal SearchDeep As Long = 999) As Collection
    ' Получает в качестве параметра путь к папке FolderPath,
    ' маску имени искомых файлов Mask (будут отобраны только файлы с такой маской/расширением)
    ' и глубину поиска SearchDeep в подпапках (если SearchDeep=1, то подпапки не просматриваются).
    ' Возвращает коллекцию, содержащую полные пути найденных файлов
    ' (применяется рекурсивный вызов процедуры GetAllFileNamesUsingFSO)

    Set FilenamesCollection = New Collection    ' создаём пустую коллекцию
    Set fso = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject
    GetAllFileNamesUsingFSO FolderPath, Mask, fso, FilenamesCollection, SearchDeep ' поиск
    Set fso = Nothing: Application.StatusBar = False    ' очистка строки состояния Excel
End Function

Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef fso, _
                                 ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
    ' перебирает все файлы и подпапки в папке FolderPath, используя объект FSO
    ' перебор папок осуществляется в том случае, если SearchDeep > 1
    ' добавляет пути найденных файлов в коллекцию FileNamesColl
    On Error Resume Next: Set curfold = fso.GetFolder(FolderPath)
    If Not curfold Is Nothing Then    ' если удалось получить доступ к папке

        ' раскомментируйте эту строку для вывода пути к просматриваемой
        ' в текущий момент папке в строку состояния Excel
        Application.StatusBar = "Поиск в папке: " & FolderPath

        For Each fil In curfold.Files    ' перебираем все файлы в папке FolderPath
            If fil.Name Like "*" & Mask Then FileNamesColl.Add fil.Path
        Next
        SearchDeep = SearchDeep - 1    ' уменьшаем глубину поиска в подпапках
        If SearchDeep Then    ' если надо искать глубже
            For Each sfol In curfold.Subfolders    ' ' перебираем все подпапки в папке FolderPath
                GetAllFileNamesUsingFSO sfol.Path, Mask, fso, FileNamesColl, SearchDeep
            Next
        End If
        Set fil = Nothing: Set curfold = Nothing    ' очищаем переменные
    End If
End Function




