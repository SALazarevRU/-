Attribute VB_Name = "mGeneralSubs"
'---------------------------------------------------------------------------------------
' Author : The_Prist(Щербаков Дмитрий)
'          Профессиональная разработка приложений для MS Office любой сложности
'          Проведение тренингов по MS Excel
'          https://www.excel-vba.ru
'          info@excel-vba.ru
'          WebMoney - R298726502453; Яндекс.Деньги - 41001332272872
' Purpose:
'---------------------------------------------------------------------------------------
Option Explicit
Const sSh_MenuBarName As String = "Список листов"
Const sSh_ListButCptn As String = "Листы"
Dim bNotEvents As Boolean
Sub Create_Sh_Menu()
    Call Del_Bar
    With Application.CommandBars.Add(sSh_MenuBarName, , , True)
        With .Controls.Add(3)
            .Caption = sSh_ListButCptn
            .OnAction = "Activate_Sh"
        End With
        .Visible = True
    End With
    Call Create_Sh_List
End Sub

Sub Del_Bar()
    On Error Resume Next
    Application.CommandBars(sSh_MenuBarName).Delete
End Sub

Sub Create_Sh_List()
    If bNotEvents Then Exit Sub
    Dim wsSh As Worksheet, li As Long
    Dim ocb
    On Error Resume Next
    Set ocb = Application.CommandBars(sSh_MenuBarName)
    If ocb Is Nothing Then
        Call Create_Sh_Menu
    End If
    On Error GoTo 0
    With Application.CommandBars(sSh_MenuBarName).Controls(sSh_ListButCptn)
        .Clear
        If lCountWorkbooks > 0 Then
            For Each wsSh In ActiveWorkbook.Sheets
                If wsSh.Visible = -1 Then
                    .AddItem wsSh.Name
                    If wsSh.index <= ActiveSheet.index Then li = li + 1
                End If
            Next wsSh
            .ListIndex = li
        Else
            .Clear
        End If
    End With
End Sub

Sub Activate_Sh()
    If lCountWorkbooks = 0 Then Exit Sub
    bNotEvents = True
    Dim sShName As String
    sShName = Application.CommandBars(sSh_MenuBarName).Controls(sSh_ListButCptn).text
    If sShName = "" Then bNotEvents = False: Exit Sub
    ActiveWorkbook.Sheets(sShName).Activate
    bNotEvents = False
End Sub

Public Function lCountWorkbooks() As Long
    Dim lCount As Long, wbBook As Workbook
    For Each wbBook In Application.Workbooks
        If wbBook.Windows(1).Visible Then lCount = lCount + 1
    Next wbBook
    lCountWorkbooks = lCount
End Function
