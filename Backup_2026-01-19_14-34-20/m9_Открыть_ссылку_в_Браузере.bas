Attribute VB_Name = "m9_Открыть_ссылку_в_Браузере"

Sub OpenWebPageAPI() 'Открыть_ссылку_в_Браузере по умолчанию (Chrome)
    Dim URL As String
    URL = "https://yandex.ru/search/?text=vba+excel....&lr=65&clid=2411726%2F/"
    ShellExecute 0, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus
End Sub





Sub SaveScreenshotAsJPG()

' Объявите переменные
Dim sImagePath As String
Dim oScreenShot As Object

' Получите путь к файлу изображения
sImagePath = "D:\OneDrive\ECXELнаOneDrive" & "\" & "Screenshot.jpg"

' Создать объект скриншота
oScreenShot = CreateObject("Screen")

' Сделайте скриншот
oScreenShot.CapturePicture sImagePath

' Сохраните скриншот в формате JPG
oScreenShot.SaveAs sImagePath, 2
End Sub




'Тоже не нашел в справке способа эмуляции нажатия кнопки Win через SendKeys.
'Но можно реализовать желаемое через Windows API, например:

'Const _
'    VK_RIGHT = &H27&, _
'    VK_LWIN = &H5B&, _
'    USE_VIRTUAL_CODES = 0&, _
'    KEY_FLAGS_NONE = 0&, _
'    KEYEVENT_KEYUP = &H2&, _
'    EXTRAS_NONE = 0&
'
'
'    Private Declare PtrSafe Sub keybd_event Lib "User32.dll" ( _
'        ByVal virtualKeyCode As Byte, _
'        ByVal hardwareScanCode As Byte, _
'        ByVal flags As Long, _
'        ByVal extraInfo As Long)

 
Sub SendExcelToRight()
    keybd_event VK_LWIN, USE_VIRTUAL_CODES, KEY_FLAGS_NONE, EXTRAS_NONE
    keybd_event VK_RIGHT, USE_VIRTUAL_CODES, KEY_FLAGS_NONE, EXTRAS_NONE
    keybd_event VK_LWIN, USE_VIRTUAL_CODES, KEYEVENT_KEYUP, EXTRAS_NONE
    keybd_event VK_RIGHT, USE_VIRTUAL_CODES, KEYEVENT_KEYUP, EXTRAS_NONE
End Sub

