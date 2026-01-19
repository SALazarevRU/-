Attribute VB_Name = "m9_—канирование_‘ Ѕ_боксы"
Option Explicit



Private Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub


'Public Sub ќбновитьЅоксы‘ Ѕ()
Sub ќбновитьЅоксы‘ Ѕ(control As IRibbonControl)
Dim sName As String
sName = "ппон‘ Ѕ"

Call ѕодсчетѕодпапкок_— Ћјƒ_‘ Ѕ

  On Error Resume Next
If Worksheets(sName) Is Nothing Then
MsgBox "Ћиста 'ппон‘ Ѕ' не существует. Exit Sub.", vbCritical
    Exit Sub  'действи€, если листа нет
Else
    'действи€, если лист есть:
   If Not gRibbon Is Nothing Then ' обновл€ю текстбоксы_‘ Ѕ
        gRibbon.InvalidateControl "editBox_—делано«а—егодн€—трок‘ Ѕ"
         gRibbon.InvalidateControl "editBox_—делано«а—егодн€ѕапок‘ Ѕ"
          gRibbon.InvalidateControl "editBox_—делано«а—егодн€‘айлов‘ Ѕ"
           gRibbon.InvalidateControl "editBox_ѕлан_‘ Ѕ"
            gRibbon.InvalidateControl "editBox_Ќомер_коробки‘ Ѕ"
             gRibbon.InvalidateControl "editBox_¬рем€—канировани€‘ Ѕ"
              gRibbon.InvalidateControl "editBox_—корость—канировани€‘ Ѕ"
               gRibbon.InvalidateControl "editBox_«апасы‘айлов‘ Ѕ"
                gRibbon.InvalidateControl "editBox_«апасыѕапок‘ Ѕ"
                gRibbon.InvalidateControl "editBox_«апасы—клад‘ Ѕ"
    End If
                        
'  On Error Resume Next
'If Worksheets(sName) Is Nothing Then
'    'действи€, если листа нет
'Else
'    'действи€, если лист есть
   End If
''≈сли лист есть, то выражение Worksheets(sName) Is Nothing ложно и управление переходит на Else.
''≈сли листа нет, то выражение Worksheets(sName) вызывает ошибку и, в буквальном соответствии
''с инструкцией On Error Resume Next управление передаетс€ следующему оператору, т.е. после Then!
End Sub


Public Sub ќбновитьЅоксы‘ Ѕ_ручник()
Dim sName As String
sName = "ппон‘ Ѕ"
  On Error Resume Next
    If Worksheets(sName) Is Nothing Then
    MsgBox "Ћиста 'ппон‘ Ѕ' не существует. Exit Sub.", vbCritical
        Exit Sub  'действи€, если листа нет
    Else
        'действи€, если лист есть:
       If Not gRibbon Is Nothing Then ' обновл€ю текстбоксы_‘ Ѕ
            gRibbon.InvalidateControl "editBox_—делано«а—егодн€—трок‘ Ѕ"
             gRibbon.InvalidateControl "editBox_—делано«а—егодн€ѕапок‘ Ѕ"
              gRibbon.InvalidateControl "editBox_—делано«а—егодн€‘айлов‘ Ѕ"
               gRibbon.InvalidateControl "editBox_ѕлан_‘ Ѕ"
                gRibbon.InvalidateControl "editBox_Ќомер_коробки‘ Ѕ"
                 gRibbon.InvalidateControl "editBox_¬рем€—канировани€‘ Ѕ"
                  gRibbon.InvalidateControl "editBox_—корость—канировани€‘ Ѕ"
                   gRibbon.InvalidateControl "editBox_«апасы‘айлов‘ Ѕ"
                    gRibbon.InvalidateControl "editBox_«апасыѕапок‘ Ѕ"
                     gRibbon.InvalidateControl "editBox_«апасы—клад‘ Ѕ"
        End If
     End If
End Sub


Sub ѕринудительно»нициализироватьЋенту()
    Dim cb As Object
    On Error GoTo ќшибка
    
    ' ѕытаемс€ получить доступ к ленте через Application
    Set cb = Application.CommandBars("Ribbon")
    If Not cb Is Nothing Then
        ' »митируем перезагрузку (это вызовет RibbonLoaded)
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        Debug.Print "ѕопытка принудительной перезагрузки ленты выполнена"
    Else
        MsgBox "Ќе удалось получить доступ к ленте", vbExclamation
    End If
    Exit Sub
    
ќшибка:
    MsgBox "ќшибка: " & Err.Description, vbCritical
End Sub

Sub ѕроверитьќбновлениеЋенты()
Application.DisplayAlerts = False
Workbooks("Ќадстройка2.xlam").ChangeFileAccess xlReadOnly, False
    Application.Wait Now + TimeValue("00:00:01")
    Workbooks("Ќадстройка2.xlam").ChangeFileAccess xlReadWrite, True
    Application.DisplayAlerts = True

End Sub


'Public Sub Ќомер оробки‘ Ѕ(editBox As IRibbonControl, ByRef text) ''это до того как номер стал забиватьс€ в эдитбокс
'On Error GoTo Instruk
'    Dim Ќом оробки‘ Ѕ  As Long
'    Ќом оробки‘ Ѕ = Worksheets("ппон‘ Ѕ").Range("D38")
'    text = "   " & Ќом оробки‘ Ѕ
'Instruk: Exit Sub
'End Sub

Public Sub ѕлан‘ Ѕ(editBox As IRibbonControl, ByRef text)
    Dim wb As Workbook
    Set wb = ActiveWorkbook 'убеждаемс€, что работаем с книгой пользовател€
    If SheetExists("ппон‘ Ѕ", wb) Then
        text = CStr(wb.Worksheets("ппон‘ Ѕ").Range("D35").Value)
    Else
        text = ""
    End If
End Sub



Public Sub ¬рем€—канировани€‘ Ѕ(editBox As IRibbonControl, ByRef text)
    Dim wb As Workbook
    Set wb = ActiveWorkbook ' работаем с книгой пользовател€

    If SheetExists("ппон‘ Ѕ", wb) Then
        text = "   " & CStr(wb.Worksheets("ппон‘ Ѕ").Range("D41").Value) & " секунд"
    Else
        text = ""
    End If
End Sub


Public Sub ¬рем€—канировани€‘ Ѕ_ ќѕ»яяяя(editBox As IRibbonControl, text As Variant)   ' ѕри заполнении в боксе —–ј«” мен€етс€ в €чейке!
    Dim editBox_¬рем€—канировани€‘ Ѕ As Variant
    editBox_¬рем€—канировани€‘ Ѕ = text
    Worksheets("ппон‘ Ѕ").Range("D41") = editBox_¬рем€—канировани€‘ Ѕ
'    Worksheets("–асширенный реестр").Range("AX1") = ¬рем€—кан
    text = "   " & editBox_¬рем€—канировани€‘ Ѕ
End Sub


Public Sub —корость—канировани€‘ Ѕ(editBox As IRibbonControl, ByRef text)
On Error GoTo Instruk
    Dim —корость—канировани€‘ Ѕ  As Variant
    —корость—канировани€‘ Ѕ = Worksheets("ппон‘ Ѕ").Range("D44")
    text = "  " & —корость—канировани€‘ Ѕ
Instruk: Exit Sub
End Sub


Public Sub —делано«а—егодн€—трок‘ Ѕ(editBox As IRibbonControl, ByRef text)
On Error GoTo Instruk
    Dim —делано«а—егодн€—трок‘ Ѕ  As Long
    —делано«а—егодн€—трок‘ Ѕ = Worksheets("ппон‘ Ѕ").Range("D47")
    text = "   " & —делано«а—егодн€—трок‘ Ѕ
Instruk: Exit Sub
End Sub

Public Sub —делано«а—егодн€ѕапок‘ Ѕ(editBox As IRibbonControl, ByRef text)
On Error GoTo Instruk
    Dim —делано«а—егодн€ѕапок‘ Ѕ  As Long
    —делано«а—егодн€ѕапок‘ Ѕ = Worksheets("ппон‘ Ѕ").Range("D48")
    text = "   " & —делано«а—егодн€ѕапок‘ Ѕ
Instruk: Exit Sub
End Sub

Public Sub —делано«а—егодн€‘айлов‘ Ѕ(editBox As IRibbonControl, ByRef text)
On Error GoTo Instruk
    Dim —делано«а—егодн€‘айлов‘ Ѕ  As Long
    —делано«а—егодн€‘айлов‘ Ѕ = Worksheets("ппон‘ Ѕ").Range("D49")
    text = "   " & —делано«а—егодн€‘айлов‘ Ѕ
Instruk: Exit Sub
End Sub

Public Sub «апасы‘айлов‘ Ѕ(editBox As IRibbonControl, ByRef text)
    Dim wb As Workbook
    Set wb = ActiveWorkbook ' работаем с книгой пользовател€

    If SheetExists("ппон‘ Ѕ", wb) Then
        text = "   " & CStr(wb.Worksheets("ппон‘ Ѕ").Range("D52").Value)
    Else
        text = ""
    End If
End Sub

'Public Sub «апасыѕапок‘ Ѕ(editBox As IRibbonControl, text As Variant)   ' «апасы ѕапок ‘ Ѕ
     'Dim «апасыѕапок‘ Ѕ  As Variant
    '«апасыѕапок‘ Ѕ = Worksheets("ппон‘ Ѕ").Range("D53")
    'text = "   " & «апасыѕапок‘ Ѕ
'End Sub


Public Sub «апасыѕапок‘ Ѕ(editBox As IRibbonControl, ByRef text)
    Dim wb As Workbook
    Set wb = ActiveWorkbook ' работаем с книгой пользовател€
    
    If SheetExists("ппон‘ Ѕ", wb) Then
        text = "   " & CStr(wb.Worksheets("ппон‘ Ѕ").Range("D53").Value)
    Else
        text = ""
    End If
End Sub

Public Sub «апасы—клад‘ Ѕ(editBox As IRibbonControl, ByRef text)
    Dim wb As Workbook
    Set wb = ActiveWorkbook ' работаем с книгой пользовател€
    
    If SheetExists("ппон‘ Ѕ", wb) Then
        text = "   " & CStr(wb.Worksheets("ппон‘ Ѕ").Range("D56").Value)
    Else
        text = ""
    End If
End Sub



'========= CALLBACK, который получает значение из EditBox ====================
Sub Ќомер оробки‘ Ѕ_Change(control As IRibbonControl, text As String)
    On Error GoTo SafeExit
    
    Dim newVal As Variant
    
    'если пользователь ввЄл не число Ч ничего не пишем
    If IsNumeric(Trim(text)) Then
        newVal = CLng(text)
    Else
        newVal = 0
    End If
    
    Worksheets("ппон‘ Ѕ").Range("D38").Value = newVal
    
SafeExit:
End Sub


'ѕроцедура, котора€ Ђзаполн€етї EditBox значением из €чейки при каждом обновлении ленты:

Sub Ќомер оробки‘ Ѕ_GetText(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = "   " & Worksheets("ппон‘ Ѕ").Range("D38").Value
End Sub

Sub ѕлан(control As IRibbonControl)
    MsgBox "0.  —делать 160 сканов." & _
    vbNewLine & "1.  «аполнить ƒинамику." & _
    vbNewLine & "2.  «аполнить ќтчет по клаймам." & _
    vbNewLine & "3.  —лить сканы на  диск Q." & _
    vbNewLine & "4.  —лить сканы в јрхив.", vbYes, "Ќу конечно же " & Application.Name & " напомнит ¬ам о планах на сегодн€: "
End Sub




'=============== проверка существовани€ листа ===============
Private Function SheetExists(shName As String, _
                             Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook  'если книга не указана Ц активна€
    On Error Resume Next
    Dim sht As Worksheet
    Set sht = wb.Worksheets(shName)
    SheetExists = Not sht Is Nothing              'True, если объект получен
    On Error GoTo 0
End Function
'============================================================














