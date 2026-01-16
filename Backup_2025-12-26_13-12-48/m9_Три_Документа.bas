Attribute VB_Name = "m9_Три_Документа"
Option Explicit

Sub Три_Документа()
    Dim Start As Date
    
    SetCursorPos 973, 500         'клик

               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
               
    Sleep (300)
    SetCursorPos 995, 582           'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("анкета")
           
           SetCursorPos 1050, 620          'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    SetCursorPos 1050, 691         'клик на добавить документ
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("паспорт")
    Sleep (300)
    SetCursorPos 1050, 740          'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300) '
    SetCursorPos 1050, 820        'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    SetCursorPos 1050, 840          'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("договор займа")
    Sleep (300)
    SetCursorPos 1050, 870          'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    SetCursorPos 1723, 956          'клик добавить 4-й документ
               mouse_event &H2, 0, 0, 0, 0
               Sleep (1500)
               mouse_event &H4, 0, 0, 0, 0
                Sleep (300)
    SetCursorPos 1050, 856         'клик на добавить документ
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    SetCursorPos 1050, 866         'клик на
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
               Sleep (300)
    Application.SendKeys ("заявление/")
    Sleep (300)
    SetCursorPos 1050, 904         'клик на заявление/акцепт договора-оферты
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0

     Sleep (300)
     SetCursorPos 1723, 253          'бегунок дернуть вверх
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
'Stop
     SetCursorPos 1350, 571          'позишн для страниц 1-го документа
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
               Sleep (300)
    Application.SendKeys ("1")
    Sleep (300)
    SetCursorPos 1420, 700          'позишн для страниц 2-го документа
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("2")
    Sleep (300)
    SetCursorPos 1330, 830          'позишн для страниц 3-го документа
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("4")
'  Stop
     
'        AppActivate ("Валидация_My_2.xlsm")  ' Активирую книгу.
'        SendKeys "{NUMLOCK}"
'           If MsgBox("Дальше?", vbYesNo) <> vbYes Then Exit Sub
End Sub

Sub Добавить_2_Дока(control As IRibbonControl)
    If MsgBox("На домашнем компе лучше не запускать эту подпрограмму, для выхода нажмите кнопку 'НЕТ'", vbYesNo) <> vbYes Then Exit Sub
    Dim Start As Date
    
    SetCursorPos 973, 500         'клик

               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
               
    Sleep (300)
    SetCursorPos 995, 582           'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("анкета")
           
           SetCursorPos 1050, 620          'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    SetCursorPos 1050, 691         'клик на добавить документ
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("паспорт")
    Sleep (300)
    SetCursorPos 1050, 740          'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300) '
    SetCursorPos 1050, 820        'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    SetCursorPos 1050, 840          'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("договор займа")
    Sleep (300)
    SetCursorPos 1050, 870          'клик
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    SetCursorPos 1723, 956          'клик добавить 4-й документ
               mouse_event &H2, 0, 0, 0, 0
               Sleep (1500)
               mouse_event &H4, 0, 0, 0, 0
                Sleep (300)
    SetCursorPos 1050, 856         'клик на добавить документ
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    SetCursorPos 1050, 866         'клик на
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
               Sleep (300)
    Application.SendKeys ("заявление/")
    Sleep (300)
    SetCursorPos 1050, 904         'клик на заявление/акцепт договора-оферты
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0

     Sleep (300)
     SetCursorPos 1723, 253          'бегунок дернуть вверх
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
'Stop
     SetCursorPos 1350, 571          'позишн для страниц 1-го документа
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
               Sleep (300)
    Application.SendKeys ("1")
    Sleep (300)
    SetCursorPos 1420, 700          'позишн для страниц 2-го документа
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("2")
    Sleep (300)
    SetCursorPos 1330, 830          'позишн для страниц 3-го документа
               mouse_event &H2, 0, 0, 0, 0
               Sleep (300)
               mouse_event &H4, 0, 0, 0, 0
    Sleep (300)
    Application.SendKeys ("4")
'  Stop
     
'        AppActivate ("Валидация_My_2.xlsm")  ' Активирую книгу.
'        SendKeys "{NUMLOCK}"
'           If MsgBox("Дальше?", vbYesNo) <> vbYes Then Exit Sub
End Sub

