Attribute VB_Name = "m9_Администратор"
Option Explicit

Sub ОткрытьРедактор(ByVal control As IRibbonControl)
    Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "s.lazarev"
'    Dim username As String, SpecifiedUserName As String: SpecifiedUserName = "Хозяин"
       username = Environ("UserName")  ' Получаем имя пользователя.
        If username = SpecifiedUserName Then
            On Error Resume Next
        '     On Error GoTo ErrHandler
               Dim Wb          As Workbook
               Set Wb = Workbooks("Надстройка2.xlam")
            
               If Wb Is Nothing Then
                   MsgBox "Файл '" & "Надстройка2.xlam" & "' не открыт.", vbExclamation
                   Exit Sub
               End If
            
               Dim vbProj      As VBIDE.VBProject
               Set vbProj = Wb.VBProject
            
               Dim vbComp      As VBIDE.VBComponent
               Set vbComp = vbProj.VBComponents("Module1")
            
               If vbComp Is Nothing Then
                   MsgBox "Модуль '" & "Module1" & "' не найден в проекте '" & "Надстройка2.xlam" & "'.", vbExclamation
                   Exit Sub
               End If
            
               Application.VBE.MainWindow.Visible = True
               Application.VBE.ActiveWindow.Project = vbProj
               vbComp.Activate
        Else
                CreateObject("WScript.Shell").Popup ("Доступа нет" & _
                vbNewLine & "Это окно сейчас закроется "), 1, "Информация о блокировке доступа к выполнению программы", 48
            Exit Sub
        End If
'ErrHandler:             'MsgBox "ТЕКСТ"
'     CreateObject("WScript.Shell").Popup ("В настоящее время доступ заблокирован." & _
'                vbNewLine & "Это окно сейчас закроется "), 1, "Информация о блокировке доступа к выполнению программы", 48
'
'    On Error GoTo 0  ' Сброс обработчика ошибок
'    Exit Sub
End Sub

Sub ОПрограмме(control As IRibbonControl)
          CreateObject("WScript.Shell").Popup ("В разработке...." & _
                vbNewLine & vbNewLine & "Это окно сейчас закроется "), 2, "Информация о программе", 48
End Sub

Sub ПроверкаОбновленияUI(control As IRibbonControl)
Dim FileName As String
    FileName = VBA.FileSystem.Dir("C:\Users\Хозяин\Desktop\Обновление_UI\Надстройка2*.xl*")
    If FileName = VBA.Constants.vbNullString Then
        MsgBox "Вы работаете с последней версией программы."
    Else
        MsgBox "Имеется новая версия панели."
    End If
    
End Sub

Sub ДоступныеНадстройки(control As IRibbonControl)
Application.CommandBars.ExecuteMso "AddInManager"
End Sub

Sub УдалитьПроцедуру(control As IRibbonControl)
Dim iProcedure As String
Dim iVBComponent As Object
Dim iStartLine As Long
Dim iCountLines As Long
Dim Killed As Boolean
 If InputBox("Введите пароль Администратора") <> "123" Then MsgBox "Неправильный пароль": Exit Sub
   iProcedure = InputBox(Prompt:="Введите имя процедуры," & _
   vbCrLf & "которую требуется удалить", title:="Удаление подпрограммы")
   If iProcedure$ = "" Then _
   MsgBox "Вы не указали имя ненужной процедуры", 48, "Ошибка": Exit Sub
   For Each iVBComponent In Workbooks("Надстройка2.xlam").VBProject.VBComponents
       With iVBComponent.CodeModule
            If .Find("Sub " & _
               iProcedure$, 1, 1, .CountOfLines, 1) = True Then
               iStartLine& = .ProcStartLine(iProcedure$, 0)
               iCountLines& = .ProcCountLines(iProcedure$, 0)
               .DeleteLines iStartLine&, iCountLines&
               Killed = True
               Exit For
            End If
       End With
   Next
   If Killed = True Then
       MsgBox "Процедура " & iProcedure$ & " удалёна!", 64, "Удаление процедуры"
   Else
       MsgBox "Процедура " & iProcedure$ & " не найден!", 48, "Удаление процедуры"
   End If
End Sub

'Sub Идентик_панели(control As IRibbonControl)   ' идентифицировать надстройку из которой выполняется код
Sub Идентик_панели()
    Dim X As Long
    Dim sTemp As String
    Dim sFileName As String
    
    Dim ИмяПроцедуры As String
    Dim ИмяМодуля As String
    Dim ВерсияПанели As String
    
    On Error GoTo ErrHandler
 
    ' Get the internal name of the add-in
    sTemp = WhoAmIToday

    For X = 1 To Application.AddIns.Count
        On Error Resume Next
        ' Attempt to call the same WhoAmI routine on each addin in the addins collection
        If Application.Run(Application.AddIns(X).Name & "!WhoAmIToday") = sTemp Then
            If Err.Number = 0 Then
                ' Found it; here's the name
                ВерсияПанели = Mid(Application.AddIns(X).Name, 1, 11)
      
            End If
        End If
    Next
    
    ИмяМодуля = ActiveWorkbook.VBProject.VBE.ActiveCodePane.CodeModule.Parent.Name
    ИмяПроцедуры = "Идентик_панели(control As IRibbonControl)"
        
     MsgBox "Текущая версия панели: " & ВерсияПанели & _
     vbNewLine & "Модуль: " & ИмяМодуля & _
     vbNewLine & "Процедура: " & ИмяПроцедуры
     
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка: " & Err.Description & ВерсияПанели & _
     vbNewLine & "Имя Модуля: " & ИмяМодуля & _
     vbNewLine & "Имя Процедуры: " & ИмяПроцедуры, vbExclamation, "Ошибка выполнения"
   On Error GoTo 0
   Exit Sub

End Sub

Sub Идентик_панели_для_бэкапа()   ' идентифицировать надстройку из которой выполняется код
    Dim X As Long
    Dim sTemp As String
    Dim sFileName As String
    
    ' Get the internal name of the add-in
    sTemp = WhoAmIToday

    For X = 1 To Application.AddIns.Count
        On Error Resume Next
        ' Attempt to call the same WhoAmI routine on each addin in the addins collection
        If Application.Run(Application.AddIns(X).Name & "!WhoAmIToday") = sTemp Then
            If Err.Number = 0 Then
                ' Found it; here's the name
                ВерсияПанели = Application.AddIns(X).Name
                
                MsgBox ВерсияПанели
            End If
        End If
    Next
End Sub

Function WhoAmIToday() As String
   WhoAmIToday = "Yes. It's ME"
End Function

'Использование планировщика задач Windows для планирования открытия книги
'https://brainbell.com/tutorials/ms-office/excel/Run_A_Macro_At_A_Set_Time.htm
'----------------------------------------------------------------------------------------------'

'Application.OnTime как запускать каждый день, в определенное время https://www.excel-vba.ru/forum/index.php?topic=4294.0&ysclid=mbqzs5kmpb571093561
Sub AutoRunMacro()
Application.OnTime TimeValue("07:30:00"), "my_Procedure"
End Sub

Sub my_Procedure()
'основной код процедуры
'после работы кода ставим опять в очередь на вызов
Application.OnTime TimeValue("07:30:00"), "my_Procedure"
End Sub

Sub Restart_этой_книги(control As IRibbonControl)
'Sub Restart_этой_книги()
Application.DisplayAlerts = False
Workbooks("Надстройка2.xlam").ChangeFileAccess xlReadOnly, False
    Application.Wait Now + TimeValue("00:00:01")
    Workbooks("Надстройка2.xlam").ChangeFileAccess xlReadWrite, True
    
'    Workbooks("03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.xlsx").ChangeFileAccess xlReadOnly, False
'    Application.Wait Now + TimeValue("00:00:01")
'    Workbooks("03122024. Расширенный реестр ГПБ к ДЦ от 26.11.2024.xlsx").ChangeFileAccess xlReadWrite, True
'    Application.DisplayAlerts = True
End Sub

Sub СохранитьВсеМодули(control As IRibbonControl)
        Dim Module As Object, newFolderPath As String, answer As String, currentTime As String, folderName As String, xFilePath As String
       
       
'        xFilePath = "E:\OneDrive\ECXELнаOneDrive\Надстройки\Бэкапы всех модулей по датам"   ' Указал путь к папке
         xFilePath = "C:\Users\s.lazarev\\Desktop\Надстройки\Бэкапы всех модулей по датам"   ' Указал путь к папке
        
        currentTime = Format(Now, "yyyy-mm-dd_hh-mm-ss")  ' Получил текущую дату и время в нужном формате
       
        folderName = "Backup_" & currentTime    ' Создал имя папки
       
        newFolderPath = xFilePath & "\" & folderName   ' Полный путь к новой папке
        
        ' Создайте папку
        On Error Resume Next ' Игнорировать ошибку, если папка уже существует
        MkDir newFolderPath
        On Error GoTo 0 ' Вернуться к нормальной обработке ошибок
        
        ' Вывожу сообщение о завершении копирования всех модулей:
        For Each Module In Workbooks("Надстройка2.xlam").VBProject.VBComponents
        Module.Export (newFolderPath & "\" & Module.Name & ".bas")
        
    Next Module
    answer = MsgBox("Все модули скопированы в: " & newFolderPath & " , открыть папку?", vbYesNo)
    
     If answer = vbYes Then
         Call ОткрытьПапкуБэкапыВсехМодулейПоДатам(Nothing)
     Else
         Exit Sub
     End If
End Sub



Sub ReadWrite()
'Workbooks("Копия Надстройка2.xlam").ChangeFileAccess xlReadOnly, False
'    Application.Wait Now + TimeValue("00:00:01")
ThisWorkbook.ChangeFileAccess Mode:=xlReadWrite
    Workbooks("Копия Надстройка2.xlam").ChangeFileAccess Mode:=xlReadWrite
End Sub

Sub DeleteMyAddIn()
Dim FSO As Object, FullName As String
AddIns("Надстройка2").Installed = False
FullName = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\Надстройка2.xlam"
AddIns("Надстройка2").Installed = False
 Application.DisplayAlerts = False
        Workbooks("Надстройка2.xlam").ChangeFileAccess xlReadOnly
        Kill Workbooks("Надстройка2.xlam").FullName
        Application.DisplayAlerts = True
End Sub

Sub ОтключитьНадстройку()
AddIns("Надстройка2").Installed = False
End Sub


Sub КопироватьФайл()
    Dim FSO As Object, SourcePath, DestinationPath As String
    Application.DisplayAlerts = False
    SourcePath = "C:\Users\Хозяин\Desktop\Обновление_UI\Надстройка2.xlam"
    DestinationPath = "C:\Users\Хозяин\AppData\Roaming\Microsoft\Надстройка2.xlam"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile SourcePath, DestinationPath, True
    Set FSO = Nothing
    
    Workbooks("Надстройка2.xlam").ChangeFileAccess xlReadOnly, False
    Application.Wait Now + TimeValue("00:00:01")
    Workbooks("Надстройка2.xlam").ChangeFileAccess xlReadWrite, True
    
    Workbooks("Лист Microsoft Excel (3).xlsx").ChangeFileAccess xlReadOnly, False
    Application.Wait Now + TimeValue("00:00:01")
    Workbooks("Лист Microsoft Excel (3).xlsx").ChangeFileAccess xlReadWrite, True
    
    Application.DisplayAlerts = True
End Sub
