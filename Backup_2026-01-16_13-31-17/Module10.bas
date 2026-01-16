Attribute VB_Name = "Module10"
Function GetFileInfo(PathName As String, FileName As String, Optional i)  ' НЕ РАБОЧИЙ 1
  ' ZVI:2011-08-29 http://www.planetaexcel.ru/forum.php?thread_id=31105  
  ' Получение свойств файла  
  '-------------------  
  ' PathName - папка  
  ' FileName - имя файла  
  ' i - номер свойства (см ниже),  
  ' если параметр i не указан, то возвращается массив свойств  
  '-------------------  
  'i = 0-Имя,1-Размер,2-Тип,3-Изменен,4-Дата создания,5-Открыт  
  'i = 6-Атрибуты,7-Состояние,8-Владелец,9-Автор,10-Заголовок  
  'i = 11-Тема,12-Категория,13-Страницы,14-Комментарий  
  '-------------------  
 Dim a, j&
 If Dir(PathName & IIf(Right(PathName, 1) <> "\", "\", "") & FileName) = "" Then Exit Function
 With CreateObject("Shell.Application").Namespace((PathName))
   If IsMissing(i) Then
     ReDim a(0 To 14)
     For j = 0 To UBound(a)
       a(j) = .GetDetailsOf(.ParseName((FileName)), j)
     Next
   Else
     a = .GetDetailsOf(.ParseName((FileName)), i)
   End If
 End With
 GetFileInfo = a
End Function

Sub ShowFileInfo() ' НЕ РАБОЧИЙ 2
 Dim x
 x = GetFileInfo("C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025", "ЛАЗАРЕВ_VII_03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.XLSX")
     If IsArray(x) Then
        Debug.Print "Имя:", x(0);  
        Debug.Print "Размер:", x(1);  
        Debug.Print "Тип:", x(2);  
        Debug.Print "Изменен:", x(3);  
        Debug.Print "Создан:", x(4);  
        Debug.Print "Открыт:", x(5);  
        Debug.Print "Аттрибут:", x(6);  
        Debug.Print "Состояние:", x(7);  
        Debug.Print "Владелец:", x(8);  
        Debug.Print "Автор:", x(9);  
        Debug.Print "Заголовок:", x(10);  
        Debug.Print "Тема:", x(11);  
        Debug.Print "Категории:", x(12)
        Debug.Print "Страницы:", x(13)
        Debug.Print "Комментарий:", x(14)
     Else  
        Debug.Print "Файл не найден"
     End If
End Sub

Sub Test_ShowAuthor() ' НЕ РАБОЧИЙ 3
 Debug.Print "Автор:", GetFileInfo("C:\Users\s.lazarev\Desktop\ГПБ_Сверка_и_сканирование_16.07.2025", "ЛАЗАРЕВ_VII_03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.XLSX", 9);  
End Sub

Sub ShowFileInfo222() ' НЕ РАБОЧИЙ ЛЕВЫЙ
Dim fs, f, s
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFile(filespec)
s = f.DateCreated
MsgBox s
End Sub


