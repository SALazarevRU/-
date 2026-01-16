Attribute VB_Name = "m9_Температура_Воздуха"

Public ТВоздуха As String
 
' Sub TVozduha()
  Sub TVozduha(control As IRibbonControl)
    '------------------------ температура воздуха в Новосибирске ------------------
    Dim sURI As String, oHttp As Object, HTMLcode
    Dim sDate$, pLeft As Integer, pRight As Integer, v As Variant
    Dim TVozduha As String
    '--------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    sURI = "https://world-weather.ru/pogoda/russia/novosibirsk/"
'    Debug.Print sURI
    
    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    
    oHttp.Open "GET", sURI, False
    oHttp.send
    HTMLcode = oHttp.responseText
'    Debug.Print HTMLcode
    
    pLeft = InStr(HTMLcode, "id=""weather-now-number"">") + 24
    pRight = InStr(pLeft, HTMLcode, "<span>")
    v = Val(Mid(HTMLcode, pLeft, pRight - pLeft)) ' rate

    If IsNumeric(v) Then
'        ActiveCell = v & "°C"
'        ActiveSheet.Range("K1") = v & "°C"
        ТВоздуха = v & "°C"
    End If
  
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "Бокс_Градусы"
    End If
    Set oHttp = Nothing
 
     Exit Sub
     
ErrHandler:
    MsgBox "Ошибка: " & Err.Description, vbExclamation, "Ошибка выполнения"
   On Error GoTo 0
   Set oHttp = Nothing
   Exit Sub
End Sub

Sub Градусы(editBox As IRibbonControl, ByRef text)
'    Dim ТВоздуха As String
'        ТВоздуха = Range("K1").Value
    text = "   " & ТВоздуха
End Sub


