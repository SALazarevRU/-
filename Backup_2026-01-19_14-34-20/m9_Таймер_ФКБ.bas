Attribute VB_Name = "m9_Таймер_ФКБ"

Public iTimer888 As Date

Sub TimerStartФКБ()
    On Error GoTo Instruk
    
    f1_ТаймерФКБ.Label2.ForeColor = RGB(237, 125, 49) ' оранж  RGB(117, 163, 255) ' лазурь
    f1_ТаймерФКБ.Label2.caption = ClaimID & "   " & ФИО
    f1_ТаймерФКБ.Label1.caption = Format(iTimer888, "n:ss")
    f1_ТаймерФКБ.Label1.ForeColor = RGB(153, 0, 204) ' белый 255 255 255
    iTimer888 = iTimer888 - TimeValue("0:00:01")
    If iTimer888 > 0 Then
        Application.OnTime Now + TimeValue("00:00:01"), "TimerStartФКБ" ' ждем секунду и подключаем процедуру TimerStart
    Else
        
        f1_ТаймерФКБ.Label1.caption = "Обработка завершена!"
        f1_ТаймерФКБ.Label2.caption = "Коробка: " & Box
            Start = Timer ' Пауза для прочтения текста в лэйбле.
                   Do While Timer < Start + 1
                       DoEvents
                   Loop
        Unload f1_ТаймерФКБ
    End If
    
Instruk:
    Exit Sub
'    MsgBox "Произошла ошибка: " & Err.Description
'   If MsgBox("Произошла ошибка: " & Err.Description _
            & vbNewLine & "Выйти из программы?", vbYesNo, "AHTUNG !!!    Ein Fehler ist aufgetreten !!!    Hitler Kaput !!!") = vbYes Then Exit Sub
    
End Sub

