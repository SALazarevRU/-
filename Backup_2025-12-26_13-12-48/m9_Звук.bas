Attribute VB_Name = "m9_Звук"
' Объявление API (в начале модуля)
Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
        Alias "PlaySoundA" (ByVal lpszName As String, _
        ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long

Const SND_SYNC = &H0          ' Ждать завершения
Const SND_ASYNC = &H1        ' Играть в фоне
Const SND_FILENAME = &H20000   ' Путь к файлу

Sub PlayWavAPI()
    Dim soundPath As String
    soundPath = "C:\Audio\looperman-3.WAV"
    
    ' Проверяем существование файла
    If Dir(soundPath) = "" Then
        MsgBox "Файл не найден: " & soundPath
        Exit Sub
    End If
    
    ' Воспроизводим синхронно (ждём окончания)
    PlaySound soundPath, 0, SND_SYNC Or SND_FILENAME
    
    ' Или асинхронно (без ожидания):
    ' PlaySound soundPath, 0, SND_ASYNC Or SND_FILENAME
End Sub

Sub PlayWavAPI_2()
    Dim soundPath As String
    soundPath = "C:\Windows\Media\ringout.wav"
    
    ' Проверяем существование файла
    If Dir(soundPath) = "" Then
        MsgBox "Файл не найден: " & soundPath
        Exit Sub
    End If
    
    ' Воспроизводим синхронно (ждём окончания)
    PlaySound soundPath, 0, SND_SYNC Or SND_FILENAME
    
    ' Или асинхронно (без ожидания):
    ' PlaySound soundPath, 0, SND_ASYNC Or SND_FILENAME
End Sub


Sub PlayWavAPI_Otklychenie()
    Dim soundPath As String
    soundPath = "C:\Users\s.lazarev\Downloads\Otklychenie.wav"
    
    ' Проверяем существование файла
    If Dir(soundPath) = "" Then
        MsgBox "Файл не найден: " & soundPath
        Exit Sub
    End If
    
    ' Воспроизводим синхронно (ждём окончания)
    PlaySound soundPath, 0, SND_SYNC Or SND_FILENAME
    
    ' Или асинхронно (без ожидания):
    ' PlaySound soundPath, 0, SND_ASYNC Or SND_FILENAME
End Sub

