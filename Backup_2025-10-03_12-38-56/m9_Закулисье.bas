Attribute VB_Name = "m9_Закулисье"
''--------------- ExcelBaby.com ---------------https://excelbaby.com/learn/contextmenus-element/

'Callback for btnFileSave onAction
Sub btnFileSave_Click(control As IRibbonControl)
    CommandBars.ExecuteMso "FileSave"
End Sub

'Callback for dmnuDemo getContent
Sub dmnuDemo_getContent(control As IRibbonControl, ByRef returnedVal)
    Dim xml As String
    Dim btnXML As String
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
          "<button id=""btnHelp"" imageMso=""Help"" label=""Справка"" onAction=""btnHelp_click""/>" & _
          "<button id=""btnFind"" imageMso=""FindDialog"" label=""Поиск"" onAction=""btnFind_click""/>" & _
          "</menu>"
    returnedVal = xml
End Sub

'Callback for btn1 onAction
Sub btnHelp_click(control As IRibbonControl)
    MsgBox "Help macro"
End Sub

'Callback for btn2 onAction
Sub btnFind_click(control As IRibbonControl)
    MsgBox "Find macro"
End Sub
