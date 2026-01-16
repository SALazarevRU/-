VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "UserForm2"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4710
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'---------------------------------------------------------------------------------------
' Решение       : Календарь
' Дата и время  : 14 января 2015  23:02
' Автор         : Night Ranger
'                 Яндекс.Деньги - 410012757639478
'                 Exingsteem@yandex.ru
'                 http://www.cyberforum.ru/vba/
' Описание      : Этот пример наглядно демонстрирует, как можно использовать календарь
'                 без подключения его к проекту, для этого нужна только форма
'                 совместимость версий любая
'
'                 В этой версии, теперь есть возможность запускать календарь от процедуры
'                 ShowCalendar, и указать там параметры SetDate и UnderRussianStandard
'                 Добавленна кнопка Ok, и форма помнит свою позицию
'---------------------------------------------------------------------------------------
Const jstart = 8, istart = 8 'Стартовые точки
Const gap = 5 'Разрыв
Const twip = 18 'Прямоугольник
Const cc = 6 'Размерность массива
Dim tt(cc, cc) As MSForms.ToggleButton, lb As MSForms.Label
Dim WithEvents fr As MSForms.Frame, WithEvents tb As MSForms.ToggleButton, WithEvents btn As MSForms.CommandButton
Attribute fr.VB_VarHelpID = -1
Attribute tb.VB_VarHelpID = -1
Attribute btn.VB_VarHelpID = -1
Dim WithEvents cbMonth As MSForms.ComboBox, WithEvents cbYear As MSForms.ComboBox
Attribute cbMonth.VB_VarHelpID = -1
Attribute cbYear.VB_VarHelpID = -1
Dim WithEvents chbx As MSForms.CheckBox, WithEvents ok As MSForms.CommandButton
Attribute chbx.VB_VarHelpID = -1
Attribute ok.VB_VarHelpID = -1
Dim iNext&, cr As Boolean, i&, j&, jj&, v, a$(), tbClick As Boolean, URStandard As Boolean

Public ThisDate As Date 'Переменная в которой храниться выбранная дата

Private Sub ok_Click()
    'Здесь могут быть дальнейшие инструкции после выбора даты
    'Например дату в удобном формате можно поместить в активную ячейку
    '----------------------------------------------------------------
    
    '
    '
    '

    ActiveCell = TextResult
    '----------------------------------------------------------------
    If chbx.Value Then Me.Hide
End Sub

Public Sub ShowCalendar( _
    Optional ByVal SetDate As Date, _
    Optional ByVal UnderRussianStandard As Boolean = 1)
    'ShowCalendar -Процедура вызова с параметрами
    'SetDate -Устанавливает возможность показа календаря c этой даты
    'UnderRussianStandard -Устанавливает возможность исправлять: 1 январь на 1 января
    If CDbl(SetDate) Then
        cr = False
        ThisDate = SetDate
        cbMonth.ListIndex = Month(ThisDate) - 1
        cbYear.text = Year(ThisDate): cr = True: Update
    End If
    URStandard = UnderRussianStandard
    Me.Show
End Sub

Private Function TextResult$()
    TextResult = FormatDateTime(ThisDate, vbLongDate)
    If URStandard Then
        TextResult = Format(ThisDate, "[$-FC19]d mmmm yyyy г.")
        
'        a = Split(TextResult)
'        If Right$(a(1), 1) Like "[йЙьЬ]" Then
'            Mid$(a(1), Len(a(1)), 1) = "я"
'        ElseIf Right$(a(1), 1) Like "[Тт]" Then a(1) = a(1) & "а"
'        End If
'        TextResult = Join(a)
    End If
End Function



Private Sub UserForm_Initialize()
    Dim maxWidth&, Width1&, jNext&
    maxWidth = twip * (cc + 1) * 2: Width1 = maxWidth \ 2: iNext = istart: jNext = jstart
    ThisDate = Date: Me.Caption = "Календарь"
    Set fr = Me.Controls.Add("Forms.Frame.1", "fr")
    Set lb = Me.Controls.Add("Forms.Label.1", "lb")
    Set cbMonth = Me.Controls.Add("Forms.ComboBox.1", "cbMonth")
    Set cbYear = Me.Controls.Add("Forms.ComboBox.1", "cbYear")
    Set btn = Me.Controls.Add("Forms.CommandButton.1", "btn")
    Set ok = Me.Controls.Add("Forms.CommandButton.1", "ok")
    Set chbx = Me.Controls.Add("Forms.CheckBox.1", "chbx")
    
    With lb: .Move jstart, istart, Width1
        .Font.Size = 15: .Font.Bold = 1
        iNext = iNext + .Height + gap
        jNext = jNext + .Width + gap
    End With
    With cbMonth: .Move jNext, istart, (Width1 - gap * 2) \ 2, lb.Height: .Style = 2
        For i = 1 To 12: .AddItem Format(DateSerial(0, i, 1), "mmmm"): Next
        jNext = jNext + .Width + gap
    End With
    With cbYear: .Move jNext, istart, (Width1 - gap * 2) \ 2, lb.Height: .Style = 2
        For i = 1899 To Year(ThisDate) + 100
            .AddItem CStr(i)
        Next
    End With
    
    iNext = lb.Top + lb.Height + gap
    
    With fr: .Move jstart, iNext, maxWidth, twip * (cc + 1)
        .Enabled = 0
        .SpecialEffect = 0
    End With
    For i = 0 To cc: For j = 0 To cc
        Set tt(j, i) = fr.Controls.Add("Forms.ToggleButton.1", "tt" & i & j)
        With tt(j, i):  .Move j * twip * 2, i * twip, twip * 2, twip: .Locked = i = 0
        .ForeColor = IIf(j >= 5, vbRed, vbBlue)
        .BackColor = IIf(i, vbButtonFace, vbScrollBars)
    End With: Next j, i
    jNext = jstart
    With ok: .Move jNext, iNext + fr.Height + gap, lb.Width, lb.Height: .Caption = "Ok"
        .AutoSize = 1: jNext = jNext + .Width + gap
    End With
    
    With btn: .Move jNext, iNext + fr.Height + gap, lb.Width, lb.Height: .Caption = "Сегодня"
        .AutoSize = 1: jNext = jNext + .Width + gap
    End With

    With chbx: .Move jNext, btn.Top, (jstart + maxWidth) - jNext
        .Caption = "Скрываться после выбора или Ok"
        .Value = GetSetting("Ms Office", "Calendar", "chbx", chbx.Value)
    End With
    

    Call btn_Click: Filling: lbUpdate

    With Me
        .Height = btn.Top + twip * 3
        .Width = jstart + maxWidth + twip
        If Application.Left > -100 Then
            .StartUpPosition = 0
            .Left = GetSetting("Ms Office", "Calendar", "Left", .Left)
            .Top = GetSetting("Ms Office", "Calendar", "Top", .Top)
            If .Left <= 0 Or .Left > (Application.Left + Application.Width - 100) Or _
            .Top <= 0 Or .Top > (Application.Top + Application.Height - 100) Then
                'Если сохраненная ранее позиция вышла за предел экрана
                .StartUpPosition = 2
            End If
        End If
    End With

End Sub

Private Sub lbUpdate()
    If cr = False Then Exit Sub
    lb.Caption = Format(ThisDate, "mmmm yyyy")
    If Split(lb.Caption)(0) <> cbMonth.text Then
        ThisDate = DateSerial(Year(ThisDate), cbMonth.ListIndex + 2, 0)
         lb.Caption = Format(ThisDate, "mmmm yyyy")
    End If
End Sub

Private Sub btn_Click()
    cr = False
    ThisDate = Date
    cbMonth.ListIndex = Month(ThisDate) - 1
    cbYear.text = Year(ThisDate): cr = True: Update
    
End Sub
Private Sub cbMonth_Click()
    If cr = False Then Exit Sub
    ThisDate = DateSerial(Year(ThisDate), cbMonth.ListIndex + 1, Day(ThisDate))
    Update
End Sub
Private Sub cbYear_Click()
     If cr = False Then Exit Sub
    ThisDate = DateSerial(cbYear.text, Month(ThisDate), Day(ThisDate)): Update
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    On Error Resume Next: Err.Clear: Set tb = tt((x - jstart) \ twip \ 2, (Y - iNext) \ twip)
    If Err = 0 Then
        With tb
            If .Enabled And .Locked = False Then
                For i = 1 To cc: For j = 0 To cc: With tt(j, i)
                    If (.Name = tb.Name) Then
                        ThisDate = DateSerial(cbYear.text, cbMonth.ListIndex + 1, .Caption)
                        .Value = 1: tbClick = 1: tb_Click: tbClick = 0 'Выбор произведен !
                    Else: .Value = 0
                    End If
    End With: Next j, i: End If: End With: End If
End Sub

Private Sub chbx_Click()
    If cr = False Then Exit Sub
    SaveSetting "Ms Office", "Calendar", "chbx", chbx.Value
End Sub

Sub Filling()
    For j = 0 To cc  'Понедельники вторники даты и тд
        With tt(j, 0): .Caption = WeekdayName(j + 1, 1, vbMonday): .Font.Bold = 1: End With
    Next: j = 0
    While Weekday(DateSerial(Year(ThisDate), Month(ThisDate), j)) <> 1: j = j - 1: Wend: jj = j
    For i = 1 To cc: For j = 0 To cc: v = DateSerial(Year(ThisDate), Month(ThisDate), jj) + 1
        With tt(j, i): .Caption = Day(v): .Enabled = Month(v) = Month(ThisDate)
            .Value = .Enabled And .Caption = Day(ThisDate)
    End With: jj = jj + 1: Next j, i
End Sub
Private Sub Update(): Call lbUpdate:  Filling: End Sub
Private Sub tb_Click(): If tbClick = False Then Exit Sub Else ok_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    With Me 'Перед закрытием запомнить позицию
        SaveSetting "Ms Office", "Calendar", "Left", .Left
        SaveSetting "Ms Office", "Calendar", "Top", .Top
    End With
End Sub

