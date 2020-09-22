Attribute VB_Name = "Module1"
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const WH_MOUSE = 7

Public Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hwnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type

Public hHook As Long

Public Function MouseProc(ByVal idHook As Long, ByVal wParam As Long, lParam As MOUSEHOOKSTRUCT) As Long
    If idHook < 0 Then
        MouseProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
    Else
        Debug.Print lParam.pt.x
        MouseProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
    End If
End Function

