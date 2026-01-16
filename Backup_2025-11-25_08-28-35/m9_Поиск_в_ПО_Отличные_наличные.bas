Attribute VB_Name = "m9_Поиск_в_ПО_Отличные_наличные"
Option Explicit
'Global ClaimID_2 As Range, ФИО_2 As Range, Box_2 As Range
'Sub ОбновитьБоксы()
Sub ОбновитьБоксы(control As IRibbonControl)
Dim sName As String
sName = "ппон"
  On Error Resume Next
If Worksheets(sName) Is Nothing Then
MsgBox "Листа 'ппон' не существует. Exit Sub.", vbCritical
    Exit Sub  'действия, если листа нет
Else
    'действия, если лист есть:
    If Not gRibbon Is Nothing Then
       gRibbon.InvalidateControl "editBox_Запросов" ' "Запросов" обновится при выполнении этой процедуры
    End If
        If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBox_Клаймов" '
        End If
            If Not gRibbon Is Nothing Then
               gRibbon.InvalidateControl "editBox_666_ОН" ' "CaptureText" обновится при выполнении этой процедуры
            End If
                If Not gRibbon Is Nothing Then
                   gRibbon.InvalidateControl "editBox_ЗапросовОсталось" ' "CaptureText" обновится при выполнении этой процедуры
                End If
                    If Not gRibbon Is Nothing Then
                       gRibbon.InvalidateControl "editBox_КлаймовОсталось" ' "CaptureText" обновится при выполнении этой процедуры
                    End If
                        If Not gRibbon Is Nothing Then
                           gRibbon.InvalidateControl "editBox_Даннные1" ' "CaptureText" обновится при выполнении этой процедуры
                        End If
                        If Not gRibbon Is Nothing Then
                           gRibbon.InvalidateControl "editBox_Даннные2" ' "CaptureText" обновится при выполнении этой процедуры
                        End If
                        If Not gRibbon Is Nothing Then
                           gRibbon.InvalidateControl "editBox_ЗапросовВЗапасе" ' "CaptureText" обновится при выполнении этой процедуры
                        End If
                        If Not gRibbon Is Nothing Then
                           gRibbon.InvalidateControl "editBox_КлаймовВЗапасе" ' "CaptureText" обновится при выполнении этой процедуры
                        End If
                        If Not gRibbon Is Nothing Then
                           gRibbon.InvalidateControl "editBox_ПланЗапросов" ' "CaptureText" обновится при выполнении этой процедуры
                        End If
                        
'  On Error Resume Next
'If Worksheets(sName) Is Nothing Then
'    'действия, если листа нет
'Else
'    'действия, если лист есть
   End If
''Если лист есть, то выражение Worksheets(sName) Is Nothing ложно и управление переходит на Else.
''Если листа нет, то выражение Worksheets(sName) вызывает ошибку и, в буквальном соответствии
''с инструкцией On Error Resume Next управление передается следующему оператору, т.е. после Then!
End Sub

Sub НетТранша(control As IRibbonControl)
    Cells(ActiveCell.Row, 13).Value = "нет транша"
    If Not gRibbon Is Nothing Then
       gRibbon.InvalidateControl "editBox_Запросов" ' "Запросов" обновится при выполнении этой процедуры
    End If
        If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBox_Клаймов" '
        End If
            If Not gRibbon Is Nothing Then
               gRibbon.InvalidateControl "editBox_666_ОН" ' "CaptureText" обновится при выполнении этой процедуры
            End If
End Sub

Sub НетРКО(control As IRibbonControl)
    Cells(ActiveCell.Row, 13).Value = "нет рко/рнко"
    If Not gRibbon Is Nothing Then
       gRibbon.InvalidateControl "editBox_Запросов" ' "Запросов" обновится при выполнении этой процедуры
    End If
        If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBox_Клаймов" '
        End If
            If Not gRibbon Is Nothing Then
               gRibbon.InvalidateControl "editBox_666_ОН" ' "CaptureText" обновится при выполнении этой процедуры
            End If
End Sub

Sub Нет(control As IRibbonControl)
    Cells(ActiveCell.Row, 13).Value = "нет"
    If Not gRibbon Is Nothing Then
       gRibbon.InvalidateControl "editBox_Запросов" ' "Запросов" обновится при выполнении этой процедуры
    End If
        If Not gRibbon Is Nothing Then
           gRibbon.InvalidateControl "editBox_Клаймов" '
        End If
            If Not gRibbon Is Nothing Then
               gRibbon.InvalidateControl "editBox_666_ОН" ' "CaptureText" обновится при выполнении этой процедуры
            End If
End Sub

 Sub Запросов(editBox As IRibbonControl, ByRef Text)
    Dim Документов  As Long
     On Error GoTo Instruk
    Документов = Workbooks("Архив. Поиск первички  НСК.xlsx").Sheets("ппон").Range("D16")
    Text = "   " & Документов
      Exit Sub
Instruk: Exit Sub
End Sub

Sub Клаймов(editBox As IRibbonControl, ByRef Text)
    Dim Клаймов  As Long
    On Error GoTo Instruk
    Клаймов = Workbooks("Архив. Поиск первички  НСК.xlsx").Sheets("ппон").Range("D15")
    Text = "   " & Клаймов
    
        If (Range("O1").Value Mod 2) = 0 Then         'MsgBox "Число в ячейке кратно 5"  'СОХРАНИЛ КНИГУ
           ActiveWorkbook.Save
        '  CreateObject("WScript.Shell").Popup "Число в ячейке кратно 5, Книгу сохранил.", 1, "Сообщение о резервном копировании файла", 48
'              f1_ФАЙЛ_СОХРАНЕН.Show 0
'                   Application.Wait Now + TimeValue("00:00:03")
'                   DoEvents
'              Unload f1_ФАЙЛ_СОХРАНЕН
                  
'              Application.StatusBar = "Число в ячейке кратно 2 - ФАЙЛ СОХРАНЕН"
'                    Application.Wait Now + TimeValue("00:00:05")
'              Application.StatusBar = False
        End If
    
    Exit Sub
Instruk: Exit Sub
End Sub

Sub CaptureText_ОН(editBox As IRibbonControl, Text As Variant) 'As Long 'для числового значения и As String для текстового
   Dim EditBoxТекст  As Variant 'для числового значения и As String для текстового
   EditBoxТекст = Text
   ActiveSheet.Range("ХХХХХХХХХХХХХХХХХХХХ1") = EditBoxТекст
End Sub

Sub ЗапросовОсталось(editBox As IRibbonControl, ByRef Text)
    Dim ДокументовОсталось  As Long
    On Error GoTo Instruk
    ДокументовОсталось = Workbooks("Архив. Поиск первички  НСК.xlsx").Worksheets("ппон").Range("D21")
    Text = "   " & ДокументовОсталось
      Exit Sub
Instruk: Exit Sub
End Sub

 Sub КлаймовОсталось(editBox As IRibbonControl, ByRef Text)
  On Error GoTo Instruk
    Dim КлаймовОсталось  As Long
    КлаймовОсталось = Workbooks("Архив. Поиск первички  НСК.xlsx").Worksheets("ппон").Range("D20")
    Text = "   " & КлаймовОсталось
Instruk: Exit Sub
End Sub

Sub ЗапросовВЗапасе(editBox As IRibbonControl, ByRef Text)
    Dim ЗапросовВЗапасе  As Long
     On Error GoTo Instruk
    ЗапросовВЗапасе = Workbooks("Архив. Поиск первички  НСК.xlsx").Worksheets("ппон").Range("D26")
    Text = "   " & ЗапросовВЗапасе
Instruk: Exit Sub
End Sub

Sub КлаймовВЗапасе(editBox As IRibbonControl, ByRef Text)
    Dim КлаймовВЗапасе  As Long
     On Error GoTo Instruk
    КлаймовВЗапасе = Workbooks("Архив. Поиск первички  НСК.xlsx").Worksheets("ппон").Range("D25")
    Text = "   " & КлаймовВЗапасе
Instruk: Exit Sub
End Sub

Sub ЭлектроннЗапросов(editBox As IRibbonControl, ByRef Text)
    Dim ЭлектроннЗапрос  As Long
     On Error GoTo Instruk
    ЭлектроннЗапрос = Workbooks("Архив. Поиск первички  НСК.xlsx").Worksheets("ппон").Range("C4")
    Text = "   " & ЭлектроннЗапрос
Instruk: Exit Sub
End Sub

Sub БумажныхЗапросов(editBox As IRibbonControl, ByRef Text)
    Dim БумажныхЗапросов  As Long
     On Error GoTo Instruk
    БумажныхЗапросов = Workbooks("Архив. Поиск первички  НСК.xlsx").Worksheets("ппон").Range("C3")
    Text = "   " & БумажныхЗапросов
Instruk: Exit Sub
End Sub
Sub ПланЗапросов(editBox As IRibbonControl, ByRef Text)
    Dim ПланЗапросов  As Long
     On Error GoTo Instruk
            ПланЗапросов = Workbooks("Архив. Поиск первички  НСК.xlsx").Worksheets("ппон").Range("D10")
            Text = "   " & ПланЗапросов
Instruk: Exit Sub
End Sub
