Attribute VB_Name = "Module7"

'Sub DataRefresh() 'сн€ть защиту
'    ActiveSheet.Unprotect "123"
'    ActiveWorkbook.RefreshAll
'    Application.OnTime Now + TimeValue("00:00:01"), "DataRefresh2"
'End Sub
'Sub DataRefresh2()
'    If Application.CommandBars.GetEnabledMso("RefreshStatus") Then
'        Application.OnTime Now + TimeValue("00:00:01"), " DataRefresh2"
'    Else
'        ActiveSheet.Protect "123", DrawingObjects:=True, Contents:=True, Scenarios:=True _
'        , AllowFiltering:=True, AllowUsingPivotTables:=True
'    End If
'End Sub
'
'Dim str1  As String, ¬ремќткрыти€Excel As Date, ff As Integer
'
'
'Sub DataRefresh666666662()
'¬ремќткрыти€Excel = Date
'           ff = FreeFile
'            Open ThisWorkbook.Path & "\Log.ini" For Output As ff
'            Write #ff, ¬ремќткрыти€Excel
'            Close ff
'End Sub
