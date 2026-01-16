Attribute VB_Name = "m9_Таймер_ГПБ"

Public iTimer As Date

Sub TimerStartГПБ()
    On Error GoTo Instruk
    
    f1_ТаймерГПБ.Label6.Caption = "Папок: " & iCountFolders22& & ", Файлов: " & iCountFiles22&
    f1_ТаймерГПБ.Label5.Caption = "Коробка " & Box_2
    f1_ТаймерГПБ.Label4.Caption = "Строк реестра: " & Dosye_2 '& " (строк за сегодня)" '    Сука, если в лэйбле не появляется значение глобальной переменной - СМОТРИ на ШИРИНУ ПОЛЯ лейбла !!!
    f1_ТаймерГПБ.Label3.Caption = "ClaimID  " & ClaimID_2
    f1_ТаймерГПБ.Label2.Caption = ФИО_2
    f1_ТаймерГПБ.Label1.Caption = Format(iTimer, "n:ss")
    iTimer = iTimer - TimeValue("0:00:01")
    If iTimer > 0 Then
        Application.OnTime Now + TimeValue("00:00:02"), "TimerStartГПБ"
    Else
        f1_ТаймерГПБ.Label1.Caption = "Обработка завершена!"
        f1_ТаймерГПБ.Label2.Caption = " "
        f1_ТаймерГПБ.Label3.Caption = " "
        f1_ТаймерГПБ.Label4.Caption = "Следующая " & "строка: " & (Dosye_2 + 1)
        f1_ТаймерГПБ.Label5.Caption = " "
        f1_ТаймерГПБ.Label6.Caption = " "
            
            Start = Timer ' Пауза для прочтения текста в лэйбле.
                   Do While Timer < Start + 3
                       DoEvents
                   Loop
        Unload f1_ТаймерГПБ
    End If
    
Instruk:
    Exit Sub
'    MsgBox "Произошла ошибка: " & Err.Description
'   If MsgBox("Произошла ошибка: " & Err.Description _
            & vbNewLine & "Выйти из программы?", vbYesNo, "AHTUNG !!!    Ein Fehler ist aufgetreten !!!    Hitler Kaput !!!") = vbYes Then Exit Sub
    
End Sub

