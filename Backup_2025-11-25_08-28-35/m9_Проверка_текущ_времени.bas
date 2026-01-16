Attribute VB_Name = "m9_Проверка_текущ_времени"
Sub CheckTime()
    Dim currentTime As Date
    Dim startTime As Date
    Dim endTime As Date
    
    currentTime = Time ' Получаем текущее время
    
    ' Устанавливаем время начала (7:00)
    startTime = TimeValue("07:00:00")
    
    ' Устанавливаем время окончания (17:00)
    endTime = TimeValue("17:00:00")
    
    ' Проверяем, находится ли текущее время в заданном промежутке
        If currentTime >= startTime And currentTime <= endTime Then
        MsgBox "Текущее время находится в промежутке с 7:00 до 17:00"
            Else
            MsgBox "Текущее время не находится в промежутке с 7:00 до 17:00"
        End If
End Sub

