Attribute VB_Name = "m9_ нига_¬_‘оновом_–ежиме"
Dim objApp As Excel.Application

 '√ќ…ƒј!!!!
Sub »зменение ниги¬‘оновом–ежиме()
    Dim Wb As Workbook
'    UserForm1.Show vbModeless
    
    Application.DisplayAlerts = False
    
    Set objApp = CreateObject("Excel.Application")
    objApp.Visible = False
    Set Wb = objApp.Workbooks.Open("C:\Users\’оз€ин\Desktop\ƒинамика 2025 Ёлектрозаводска€.xlsx")
    
    Wb.Worksheets("Ћист1").Range("A1:CF90000").Value = 2000
    
    Wb.Close SaveChanges:=True
    objApp.Quit
    Set objApp = Nothing
    
'    Unload UserForm1
 
    Application.DisplayAlerts = True
End Sub


Sub –абота¬‘оновом–ежиме()
    ' ќткрытие книги в фоновом режиме (Visible = False):
    Dim Wb As Workbook
     UserForm1.Show vbModeless
    Application.DisplayAlerts = False
    Set Wb = Workbooks.Open("C:\Users\’оз€ин\Desktop\ƒинамика 2025 Ёлектрозаводска€.xlsx", False)
    ' ¬ыполнение операций с книгой, например:
    Worksheets("Ћист1").Range("A1:CF90000").Value = 2000
    ' —охранение книги:
    Wb.Save
    ' «акрытие книги без отображени€ на экране:
    Wb.Close False  ' (второй параметр Ч SaveChanges, False означает Ђне сохран€ть изменени€ перед закрытиемї)
    Set Wb = Nothing
    Unload UserForm1
    Application.DisplayAlerts = True
End Sub

Sub ќткрытьјвторизаци€_2()
    ' ќткрытие книги в фоновом режиме (Visible = False):
    Dim Wb As Workbook
     
    Application.DisplayAlerts = False
    Set Wb = Workbooks.Open("C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\јвторизаци€.csv", False)
    ' ¬ыполнение операций с книгой, например:
    Ћогин = Worksheets("јвторизаци€").Range("B2").Value
    ѕарольѕќ = Worksheets("јвторизаци€").Range("B3").Value
    
    ' «акрытие книги без отображени€ на экране:
    Wb.Close False  ' (второй параметр Ч SaveChanges, False означает Ђне сохран€ть изменени€ перед закрытиемї)
    Set Wb = Nothing
    
    Application.DisplayAlerts = True
End Sub

