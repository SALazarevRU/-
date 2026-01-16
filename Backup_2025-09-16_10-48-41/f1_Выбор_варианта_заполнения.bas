VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f1_Выбор_варианта_заполнения 
   Caption         =   "©  Выбор варианта заполнения строки™"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9630.001
   OleObjectBlob   =   "f1_Выбор_варианта_заполнения.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f1_Выбор_варианта_заполнения"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub CommandButton1_Click()
   CommandButton1.BackColor = RGB(204, 58, 0)
 Set ФИО_2 = ActiveCell
'    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    
        If Cells(ActiveCell.Row, 34).Value = "" Then
            Cells(ActiveCell.Row, 34).Value = "расчет выписка"
        Else
            Cells(ActiveCell.Row, 34).Value = Cells(ActiveCell.Row, 34).Value & ", " & "расчет выписка"
            Cells(ActiveCell.Row, 43).Value = "В ячейку столбца AH были добавлены данные"
            Cells(ActiveCell.Row, 43).Interior.Color = RGB(255, 125, 125)
         End If
    
'    Cells(ActiveCell.Row, 38).Value = Box_2
    
        If Cells(ActiveCell.Row, 38).Value = "" Then
           Cells(ActiveCell.Row, 38).Value = Box_2
        Else
          MsgBox "Внимание! В ячейке уже есть номер коробки! "
        End If
        
        If Cells(ActiveCell.Row, 41).Value = "" Then
           Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
           Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
          Else
                Cells(ActiveCell.Row, 45).Value = Cells(ActiveCell.Row, 41).Value
                Cells(ActiveCell.Row, 41).ClearContents
                Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
                Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
        End If
        
        If Cells(ActiveCell.Row, 42).Value = "" Then
           Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
           Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"
          Else
                Cells(ActiveCell.Row, 46).Value = Cells(ActiveCell.Row, 42).Value
                Cells(ActiveCell.Row, 46).NumberFormat = "dd.mm.yyyy hh:mm:ss"
                Cells(ActiveCell.Row, 42).ClearContents
                Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
                Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"
        End If
        
    End If
    Unload Me
    Call СканингГПБ_Часть3 ' эта часть для перееименовки скана в "расчет выписка"

End Sub

Private Sub CommandButton11_Click()
    Set ФИО_2 = ActiveCell
    '    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate '  будет ошибка 1004
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 33).Value = "копия испол.надписи"
'    Cells(ActiveCell.Row, 37).Value = Box_2
    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
    Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"

    End If
    Unload Me
    Call СканингГПБ_Часть2
End Sub

Private Sub CommandButton12_Click()  'КОПИЯ ИСПОЛ ЛИСТА
    Set ФИО_2 = ActiveCell
    '    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Select '  будет ошибка 1004
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 33).Value = "копия ИЛ"
'    Cells(ActiveCell.Row, 37).Value = Box_2
    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
    Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    End If
    Unload Me
    Call СканингГПБ_Часть2
End Sub

Private Sub CommandButton13_Click()
Set ФИО_2 = ActiveCell
'    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
'    Set ClaimID_2 = ActiveCell
'    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
       Cells(ActiveCell.Row, 35).Value = "информация ИД "
    End If
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton14_Click()
    Set ФИО_2 = ActiveCell
        If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
           Cells(ActiveCell.Row, 37).Value = Box_2
        End If
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton15_Click()
    Set ФИО_2 = ActiveCell
    Set Box_2 = Range("AP1")
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
            If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
        End If
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton16_Click()
Set ФИО_2 = ActiveCell
        If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
           Cells(ActiveCell.Row, 33).Value = "копия пост.ФССП"
        End If
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton17_Click()
Unload Me
End Sub

Private Sub CommandButton18_Click()
' Set ФИО_2 = ActiveCell
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
     If Cells(ActiveCell.Row, 33).Value = "" Then
        Cells(ActiveCell.Row, 33).Value = "копия СП"
        Cells(ActiveCell.Row, 40).Value = "копия ИД"
     Else
      Cells(ActiveCell.Row, 33).Value = Cells(ActiveCell.Row, 33).Value & ", " & "копия СП"
      Cells(ActiveCell.Row, 40).Value = "копия ИД"
    End If
    End If
    
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton19_Click()
Set ФИО_2 = ActiveCell
        If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
           Cells(ActiveCell.Row, 33).Value = "копия ИЛ"
        End If
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton2_Click()
    Set ФИО_2 = ActiveCell
   '    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 32).Value = "оригинал"
    Cells(ActiveCell.Row, 33).Value = "копия СП"
'    Cells(ActiveCell.Row, 37).Value = Box_2
    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
    Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"

    End If
    Unload Me
    Call СканингГПБ_Часть2
End Sub

Private Sub CommandButton20_Click()
If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
     If Cells(ActiveCell.Row, 33).Value = "" Then
        Cells(ActiveCell.Row, 33).Value = "копия испол.надписи"
     Else
      Cells(ActiveCell.Row, 33).Value = Cells(ActiveCell.Row, 35).Value & ", " & "копия испол.надписи"
    End If
    End If
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton21_Click()
    Set ФИО_2 = ActiveCell
    '    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")

'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 32).Value = "оригинал"
'    Cells(ActiveCell.Row, 37).Value = Box_2
    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
        If Cells(ActiveCell.Row, 41).Value = "" Then
           Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
           Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
        Else
                Cells(ActiveCell.Row, 45).Value = Cells(ActiveCell.Row, 41).Value
                Cells(ActiveCell.Row, 41).ClearContents
                Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
                Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
        End If
        
        If Cells(ActiveCell.Row, 42).Value = "" Then
           Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
           Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"
        Else
                Cells(ActiveCell.Row, 46).Value = Cells(ActiveCell.Row, 42).Value
                Cells(ActiveCell.Row, 46).NumberFormat = "dd.mm.yyyy hh:mm:ss"
                Cells(ActiveCell.Row, 42).ClearContents
                Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
                Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"
        End If

    End If
    Unload Me
    Call СканингГПБ_Часть2
End Sub

Private Sub CommandButton22_Click()
'CommandButton21.BackColor = RGB(204, 58, 0)
 Set ФИО_2 = ActiveCell
'    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 34).Value = "расчет"
    Cells(ActiveCell.Row, 38).Value = Box_2
'    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
    Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"

    End If
    Unload Me
    Call СканингГПБ_Часть4 ' эта часть для перееименовки скана в "расчет"
End Sub

Private Sub CommandButton23_Click()
Set ФИО_2 = ActiveCell
'    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 34).Value = "выписка"
    Cells(ActiveCell.Row, 38).Value = Box_2
'    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
    Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"

    End If
    Unload Me
    Call СканингГПБ_Часть5 ' эта часть для перееименовки скана в "выписка"
End Sub

Private Sub CommandButton24_Click()
 Set ФИО_2 = ActiveCell
   '    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 32).Value = "копия"
    Cells(ActiveCell.Row, 33).Value = "копия СП"
'    Cells(ActiveCell.Row, 37).Value = Box_2
    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
    Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"

    End If
    Unload Me
    Call СканингГПБ_Часть2
End Sub

Private Sub CommandButton25_Click()
' Set ФИО_2 = ActiveCell    .ClearContents
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
      If Cells(ActiveCell.Row, 33).Value = "" Then
         Cells(ActiveCell.Row, 33).Value = "копия СП"
         Cells(ActiveCell.Row, 40).ClearContents
      Else
         Cells(ActiveCell.Row, 33).Value = Cells(ActiveCell.Row, 33).Value & ", " & "копия СП"
         Cells(ActiveCell.Row, 40).ClearContents
      End If
    End If
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton27_Click()
  If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
     If Cells(ActiveCell.Row, 33).Value = "" Then
        Cells(ActiveCell.Row, 33).Value = "копия ИЛ"
        Cells(ActiveCell.Row, 40).Value = "копия ИД"
     Else
      Cells(ActiveCell.Row, 33).Value = Cells(ActiveCell.Row, 33).Value & ", " & "копия ИЛ"
      Cells(ActiveCell.Row, 40).Value = "копия ИД"
    End If
    End If
    
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton3_Click()
    Set ФИО_2 = ActiveCell
'    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 33).Value = "копия СП"
'    Cells(ActiveCell.Row, 37).Value = Box_2
    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
    Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"

    End If
    Unload Me
    Call СканингГПБ_Часть2
End Sub

Private Sub CommandButton4_Click()
  Dim DataDogovora
        Dim title
      
        Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
        title = "Изменить дату договора займа (в Комментарии)"
            DataDogovora = InputBox("Введите дату договора," _
            & vbNewLine & "" _
            & vbNewLine & "", title)
            ActiveCell.Offset(0, 33).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
            ActiveCell.Value = "дата договора от " & DataDogovora
            Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
            '    If MsgBox("СКАНИРОВАТЬ?", vbYesNo) <> vbYes Then Exit Sub
            
End Sub

Private Sub CommandButton5_Click()
  Set ФИО_2 = ActiveCell
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
     If Cells(ActiveCell.Row, 35).Value = "" Then
        Cells(ActiveCell.Row, 35).Value = "нет договора"
     Else
      Cells(ActiveCell.Row, 35).Value = Cells(ActiveCell.Row, 35).Value & ", " & "нет договора"
    End If
    End If
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
End Sub

Private Sub CommandButton6_Click()
    Set ФИО_2 = ActiveCell
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 35).Value = Cells(ActiveCell.Row, 35).Value & ", " & "без подписи"
    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки
    End If
End Sub

Private Sub CommandButton7_Click()
    Set ФИО_2 = ActiveCell
   '    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1")
'    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 32).Value = "копия"
    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
     
    Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
    Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
    Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
    Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"

    End If
    Unload Me
    Call СканингГПБ_Часть2
End Sub

Private Sub CommandButton8_Click() 'НИЧЕГО
    Set ФИО_2 = ActiveCell
'    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
    Set ClaimID_2 = ActiveCell.Offset(0, 22)
    Set Box_2 = Range("AP1") '    без бокс2 не работае..
'    Cells(ActiveCell.Row, 2).Select ' переместиться во вторую ячейку текущей строки

    Cells(ActiveCell.Row, 39).Value = Range("AQ1").Value ' Номер коробки с копиями КД
    
    If Cells(ActiveCell.Row, 40).Value = "" Then
       Cells(ActiveCell.Row, 40).Value = "копия КД"
    Else
      Cells(ActiveCell.Row, 40).Value = Cells(ActiveCell.Row, 40).Value & " + " & "копия КД"
    End If
      
    If Cells(ActiveCell.Row, 37).Value = "" Then: Cells(ActiveCell.Row, 37).Value = Box_2
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
     If Cells(ActiveCell.Row, 41).Value = "" Then
           Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
           Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
          Else
                Cells(ActiveCell.Row, 45).Value = Cells(ActiveCell.Row, 41).Value
                Cells(ActiveCell.Row, 41).ClearContents
                Cells(ActiveCell.Row, 41).Value = Date ' Именно Date, а не Now!
                Cells(ActiveCell.Row, 41).NumberFormat = "dd.MM.yyyy"
        End If
        
        If Cells(ActiveCell.Row, 42).Value = "" Then
           Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
           Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"
          Else
                Cells(ActiveCell.Row, 46).Value = Cells(ActiveCell.Row, 42).Value
                Cells(ActiveCell.Row, 46).NumberFormat = "dd.mm.yyyy hh:mm:ss"
                Cells(ActiveCell.Row, 42).ClearContents
                Cells(ActiveCell.Row, 42).Value = Now ' Именно Now, а не Date!
                Cells(ActiveCell.Row, 42).NumberFormat = "dd.mm.yyyy hh:mm:ss"
        End If
    End If
    
    Unload Me
    Call СканингГПБ_Часть2
End Sub

Private Sub CommandButton9_Click()
    Set ФИО_2 = ActiveCell
    ActiveCell.Offset(0, 22).Activate 'Перехожу на 22 ячейки правее активной и активирую ее.
'    Set ClaimID_2 = ActiveCell
    Set Box_2 = Range("AP1")
    ActiveCell.Offset(0, -22).Activate
    If Not Intersect(ActiveCell, Range("B2:B5680")) Is Nothing Then
    Cells(ActiveCell.Row, 35).Value = "Копия в кор " & Box_2
    End If
End Sub

Private Sub UserForm_Initialize()

    Me.StartUpPosition = 0
    Me.Top = 290 + Application.Top
    Me.Left = 460 + Application.Left
End Sub
