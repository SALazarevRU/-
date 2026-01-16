Attribute VB_Name = "m9_ќткрыть_ѕапки"
'Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub ќткрытьѕапку—лива—кановќ“Ћ»„ »_ЌаQ(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "Q:\LP2\задача 51677\Ќ— "
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапку—лива—канов√ѕЅЌаQ(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "Q:\LP2\–езультаты сверки портфелей с августа 2020\√азпром √ѕЅ 2\сканы 01.08"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапкујрхивЋазарев(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "Q:\Corporative\ќ—¬\ќЅћ≈ЌЌ» \‘абрика\отдел архива длительного хранени€\архив\Ћазарев\10.04.2025"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапку—каны()
    ShellExecute hwnd, "open", "C:\Users\s.lazarev\Desktop\2_Ѕыстроденьги_сканинг\— јЌџ_в работе\", nil, nil, SW_SHOWNORMAL
End Sub

'‘ункцию ShellExecute можно использовать,например, дл€ открыти€ корневого каталога диска —:

Sub ќткрытьѕапку—каны_на_Q()
    ShellExecute hwnd, "open", "Q:\LP2\–езультаты сверки портфелей с августа 2020\Ѕыстроденьги ‘ Ѕ\сканы на бумаге\", nil, nil, SW_SHOWNORMAL
End Sub

Private Sub ќткрытьѕапкуќтчетѕо лаймамЌаQ(control As IRibbonControl)
  Dim myFolder As String
    myFolder = "Q:\Corporative\ќ—¬\ќЅћ≈ЌЌ» \‘абрика\отдел архива длительного хранени€\архив\ƒинамика\ќ“„≈“ по клаймам"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапку—канировани€‘ Ѕ(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "C:\Users\s.lazarev\Desktop\2_Ѕыстроденьги_сканинг\— јЌџ_в работе"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапку—каны‘ ЅнаQ(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "Q:\LP2\–езультаты сверки портфелей с августа 2020\Ѕыстроденьги ‘ Ѕ\сканы на бумаге"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапкуќтчет‘абрика(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "Q:\Corporative\ќ—¬\ќЅћ≈ЌЌ» \‘абрика\ƒл€ руководителей\ќперационна€ отчетность фабрики\2025"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапку—лил—каны‘ ЅнаD(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "D:\2_Ѕыстроденьги_сканинг"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапку–абочий—тол(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "C:\Users\s.lazarev\Desktop"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub ќткрытьѕапкуƒинамика(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "Q:\Corporative\ќ—¬\ќЅћ≈ЌЌ» \‘абрика\отдел архива длительного хранени€\архив\ƒинамика"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Sub Open_Folder_ƒинамика() ' работает
    Dim fdName$
    Path = "Q:\Corporative\ќ—¬\ќЅћ≈ЌЌ» \‘абрика\отдел архива длительного хранени€\архив\"
    If Dir(Path & "*.*") = "" Then
        MsgBox "ѕуть:" & vbLf & Path & vbLf & "не найден(((", vbCritical, "јварийное завешение работы!": Exit Sub
    End If
    fdName = Dir(Path & "ƒинамика*.*", vbDirectory)
    If fdName <> "" Then
        Call shell("explorer.exe " & Path & fdName, vbNormalFocus)
    Else
        MsgBox "Ќет папки ƒинамика по указанному пути ", , Path
        End If
End Sub


Sub CallќткрытьѕапкуЅэкапыѕоƒатам() 'ƒом
    Call ќткрытьѕапкуЅэкапы¬сехћодулейѕоƒатам(Nothing)
End Sub
Public Sub ќткрытьѕапкуЅэкапыѕоƒатам(control As IRibbonControl) 'ƒом
    Dim myFolder As String
    myFolder = "C:\Users\s.lazarev\Desktop\ЅЁ јѕџ" & "\"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus   'Ќќ–ћјЋ№Ќќ ќ“ –џ¬ј≈“.
End Sub

Public Sub ќткрытьѕапкуЅэкапы¬сехћодулейѕоƒатам(control As IRibbonControl) 'ƒом
    Dim myFolder As String
'    myFolder = "E:\OneDrive\ECXELнаOneDrive\Ќадстройки\Ѕэкапы всех модулей по датам" & "\"
    myFolder = "C:\Users\s.lazarev\Desktop\Ќадстройки\Ѕэкапы всех модулей по датам"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub

Public Sub ќткрытьѕапкуЅэкапы(control As IRibbonControl)
    Dim myFolder As String
    myFolder = "C:\Users\s.lazarev\Desktop\ЅЁ јѕџ" & "\"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub
  
Sub ќткрытьѕапкуЅэкапЌадстройки2_на_≈(control As IRibbonControl) 'ƒом
    Dim myFolder As String
    myFolder = "E:\OneDrive\ECXELнаOneDrive\Ќадстройки\Backap_AddIns"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub


Sub ќткрытьѕапкуTemp(control As IRibbonControl) 'ƒом
    Dim myFolder As String
        On Error GoTo ErrHandler
    myFolder = "C:\Users\s.lazarev\AppData\Local\Temp"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
    Exit Sub
ErrHandler:
    MsgBox "ќшибка: " & Err.Description, vbExclamation, "ќшибка выполнени€"
   On Error GoTo 0
   Exit Sub
End Sub

Sub ќткрытьѕапкујддоны(control As IRibbonControl) 'ƒом
    Dim myFolder As String
        On Error GoTo ErrHandler
'    myFolder = "C:\Users\’оз€ин\AppData\Roaming\Microsoft\AddIns"
    myFolder = "C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
    Exit Sub
ErrHandler:
    MsgBox "ќшибка: " & Err.Description, vbExclamation, "ќшибка выполнени€"
   On Error GoTo 0
   Exit Sub
End Sub

Sub ќткрытьѕапкуЅэкапы–абоч нигЌајрхиве(control As IRibbonControl) 'ƒом
    Dim myFolder As String
    myFolder = "Q:\Corporative\ќ—¬\ќЅћ≈ЌЌ» \‘абрика\отдел архива длительного хранени€\архив\Ћазарев\Backups\–абочие книги"
    shell "explorer.exe """ & myFolder & """", vbNormalFocus
End Sub


