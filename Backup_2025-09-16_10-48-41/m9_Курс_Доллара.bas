Attribute VB_Name = "m9_Курс_Доллара"
Dim dollarRate As Double
Sub ПолучитьКурсДоллара(control As IRibbonControl)
    Dim xmDoc As Object
    Dim dollarNode As Object
    
    Dim date_req As String
    
    ' Создание объекта XML
    Set xmDoc = CreateObject("msxml2.DOMDocument")
    
    ' Загрузка данных с сайта ЦБР
    xmDoc.async = False
    xmDoc.Load ("http://www.cbr.ru/scripts/XML_daily.asp")
    
    ' Проверяем, загружен ли документ
    If xmDoc.parseError <> 0 Then
        MsgBox "Ошибка при загрузке данных: " & xmDoc.parseError.reason, vbCritical
        Exit Sub
    End If
    
    ' Получаем узел с курсом доллара
    Set dollarNode = xmDoc.SelectSingleNode("*/Valute[CharCode='USD']")
    
    ' Проверяем, найден ли узел
    If Not dollarNode Is Nothing Then
        ' Получаем курс доллара
        dollarRate = CDbl(dollarNode.ChildNodes(4).Text) / Val(dollarNode.ChildNodes(2).Text)
        
        ' Записываем курс в ячейку A1
'        Worksheets("Лист1").Range("K1").Value = dollarRate
       
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "editBox_Dollar"
    End If
        ' Уведомление об успешном завершении
'        MsgBox "Курс доллара: " & dollarRate, vbInformation
    Else
        MsgBox "Курс доллара не найден!", vbExclamation
    End If
    
    ' Освобождение объекта
    Set xmDoc = Nothing
End Sub

Sub Деньги(editBox As IRibbonControl, ByRef Text)
'    Dim ТВоздуха As String
'        ТВоздуха = Range("K1").Value
    Text = "   " & dollarRate
End Sub
