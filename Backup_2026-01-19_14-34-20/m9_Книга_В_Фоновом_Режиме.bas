Attribute VB_Name = "m9_ нига_¬_‘оновом_–ежиме"
Dim objApp As Excel.Application

 '√ќ…ƒј!!!!
Sub »зменение ниги¬‘оновом–ежиме()
    Dim wb As Workbook
'    UserForm1.Show vbModeless
    
    Application.DisplayAlerts = False
    
    Set objApp = CreateObject("Excel.Application")
    objApp.Visible = False
    Set wb = objApp.Workbooks.Open("C:\Users\’оз€ин\Desktop\ƒинамика 2025 Ёлектрозаводска€.xlsx")
    
    wb.Worksheets("Ћист1").Range("A1:CF90000").Value = 2000
    
    wb.Close SaveChanges:=True
    objApp.Quit
    Set objApp = Nothing
    
'    Unload UserForm1
 
    Application.DisplayAlerts = True
End Sub


Sub –абота¬‘оновом–ежиме()
    ' ќткрытие книги в фоновом режиме (Visible = False):
    Dim wb As Workbook
     UserForm1.Show vbModeless
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open("C:\Users\’оз€ин\Desktop\ƒинамика 2025 Ёлектрозаводска€.xlsx", False)
    ' ¬ыполнение операций с книгой, например:
    Worksheets("Ћист1").Range("A1:CF90000").Value = 2000
    ' —охранение книги:
    wb.Save
    ' «акрытие книги без отображени€ на экране:
    wb.Close False  ' (второй параметр Ч SaveChanges, False означает Ђне сохран€ть изменени€ перед закрытиемї)
    Set wb = Nothing
    Unload UserForm1
    Application.DisplayAlerts = True
End Sub

Sub ќткрытьјвторизаци€_2()
    ' ќткрытие книги в фоновом режиме (Visible = False):
    Dim wb As Workbook
     
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open("C:\Users\s.lazarev\AppData\Roaming\Microsoft\AddIns\јвторизаци€.csv", False)
    ' ¬ыполнение операций с книгой, например:
    Ћогин = Worksheets("јвторизаци€").Range("B2").Value
    ѕарольѕќ = Worksheets("јвторизаци€").Range("B3").Value
    
    ' «акрытие книги без отображени€ на экране:
    wb.Close False  ' (второй параметр Ч SaveChanges, False означает Ђне сохран€ть изменени€ перед закрытиемї)
    Set wb = Nothing
    
    Application.DisplayAlerts = True
End Sub

