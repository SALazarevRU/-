Attribute VB_Name = "m_Все_Declare_Референсы"
Public Declare PtrSafe Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long ' для очистки корзины
Public Declare PtrSafe Function MessageBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" ()
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)  ' для функции Sleep

Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile _
                         As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Public Declare PtrSafe Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszfile As String, ByVal ucommand As Long, ByVal dwdata As Long) As Long

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr

Public Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Const SWP_NOZORDER As Long = &H4
Const SWP_NOACTIVATE As Long = &H10



Public Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SETDESKWALLPAPER = 20

'Ахтунг! Если над функциями будут расположены прцедуры- возникнет ошибка!

