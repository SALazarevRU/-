Attribute VB_Name = "m9_Бэкапы"
'Public Sub CallБэкапЭтойКниги()
'    Call БэкапЭтойКниги(Nothing)
'End Sub

Public Sub ВсеБэкапы(control As IRibbonControl)
    Call БэкапЭтойКниги(Nothing)
    Call БэкапДинамики_2025(Nothing)
    Call БэкапНадстрВалПД(Nothing)
    
    Const lSeconds As Long = 3
    MessageBoxTimeOut 0, "БЭКАПЫ ВЫПОЛНЕНЫ." & _
     vbNewLine & " " & _
     vbNewLine & "Сообщение закроется через 3 секунды", "Отчёт о завершении резервного копирования файлов", _
   vbInformation + vbOKOnly, 0&, lSeconds * 1000
End Sub

Public Sub БэкапЭтойКниги(control As IRibbonControl)
    Dim sFolderPath As String
    Dim sFileName As String
    Dim sFileExt As String
    Dim sNewFileName As String
    
    Dim sDateTimeStamp As String
    sDateTimeStamp = VBA.Format(VBA.Now, " yyyy-mm-dd  HH-MM-SS")

    sFolderPath = "C:\Users\s.lazarev\Desktop\БЭКАПЫ" & "\"
    
    sFileName = ActiveWorkbook.Name
'    MsgBox sFileName, vbInformation
    sFileExt = VBA.Mid(sFileName, VBA.InStrRev(sFileName, ".", , vbTextCompare))
    sNewFileName = VBA.Replace(sFileName, sFileExt, "", , , vbTextCompare)
    
    sNewFileName = sFolderPath & sNewFileName & " (Backup) " & sDateTimeStamp & sFileExt
    
   Call БэкапЭтойКниги_На_Q
   
        Application.DisplayAlerts = False
    ActiveWorkbook.SaveCopyAs sNewFileName
'    MsgBox "Бэкап создан", vbInformation
        Application.DisplayAlerts = True
End Sub

Public Sub БэкапЭтойКниги_На_Q()
    Dim sFolderPath As String
    Dim sFileName As String
    Dim sFileExt As String
    Dim sNewFileName As String
    
    Dim sDateTimeStamp As String
    sDateTimeStamp = VBA.Format(VBA.Now, " yyyy-mm-dd  HH-MM-SS")

    sFolderPath = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Лазарев\Backups\Рабочие книги" & "\"
    sFileName = ActiveWorkbook.Name
    sFileExt = VBA.Mid(sFileName, VBA.InStrRev(sFileName, ".", , vbTextCompare))
    sNewFileName = VBA.Replace(sFileName, sFileExt, "", , , vbTextCompare)
    
    sNewFileName = sFolderPath & sNewFileName & " (Backup) " & sDateTimeStamp & sFileExt
    
        Application.DisplayAlerts = False
    ActiveWorkbook.SaveCopyAs sNewFileName
'    MsgBox "Бэкап создан", vbInformation
        Application.DisplayAlerts = True
End Sub

Sub SaveFile_ВАЛ_WithTimeStamp()
    Dim FileName As String
    Dim sFolderPath As String
    Dim sFileName As String
    Dim sFileExt As String
    Dim sNewFileName As String
    
    FileName = "C:\Users\Хозяин\Desktop\Валидация_My_2.xlsm"
    
    Dim sDateTimeStamp As String
    sDateTimeStamp = VBA.Format(VBA.Now, " от dd_mmmm_yyyy  HH-MM-SS")

    sFolderPath = "C:\Users\Хозяин\Desktop\Бэкапы" & "\"
    sFileName = "Валидация_My_2.xlsm"
    
    sFileExt = VBA.Mid(sFileName, VBA.InStrRev(sFileName, ".", , vbTextCompare))
    sNewFileName = VBA.Replace(sFileName, sFileExt, "", , , vbTextCompare)
    
    sNewFileName = sFolderPath & sNewFileName & " (Backup) " & sDateTimeStamp & sFileExt
    
'    Workbook(sNewFileName).Save
Workbooks("Валидация_My_2.xlsm").SaveAs FileName:=sNewFileName
End Sub

Sub БэкапДинамики_2025(control As IRibbonControl)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    SourceFile = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Динамика\Динамика 2025 Электрозаводская.xlsx"
    destFolder = "C:\Users\s.lazarev\Documents\BackupДинамика\" & "Динамика 2025 Электрозаводская " & " (Backup) " & Format(Now(), "yyyy-mm-dd hh-mm-ss") & ".xlsx"
    FSO.CopyFile SourceFile, destFolder
End Sub

Sub БэкапНадстроки(control As IRibbonControl)
MsgBox "Бэкап Надстройки2 на ЭТОМ ПК будет создан лишь на диске Q \архив\Лазарев\Backups\Backap_AddIns", vbInformation
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    SourceFile = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\Надстройка2.xlam"
'    destFolder = "E:\OneDrive\ECXELнаOneDrive\Надстройки\Backap_AddIns\" & "Надстройка2" & " (Backup) " & Format(Now(), "yyyy-mm-dd hh-mm-ss") & " .xlam"
'    FSO.CopyFile SourceFile, destFolder
    Call БэкапНадстроки_На_Q
End Sub
Sub БэкапНадстроки_На_Q()
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    SourceFile = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\Надстройка2.xlam"
    destFolder = "Q:\Corporative\ОСВ\ОБМЕННИК\Фабрика\отдел архива длительного хранения\архив\Лазарев\Backups\Backap_AddIns\" & "Надстройка2" & " (Backup) " & Format(Now(), "yyyy-mm-dd hh-mm-ss") & " .xlam"
    FSO.CopyFile SourceFile, destFolder
    MsgBox "Бэкап Надстройки2 создан на диске Q \архив\Лазарев\Backups\Backap_AddIns", vbInformation
End Sub




