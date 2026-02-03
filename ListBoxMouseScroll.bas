Attribute VB_Name = "ListBoxMouseScroll"
Option Explicit
   
Private Type POINTAPI
     x As Long
     y As Long
End Type
   
Private Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        mouseData As Long
        flags As Long
        time As Long
        dwExtraInfo As LongPtr
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" _
                         Alias "FindWindowA" ( _
                                 ByVal lpClassName As String, _
                                 ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32.dll" _
                         Alias "GetWindowLongPtrA" ( _
                                 ByVal hwnd As LongPtr, _
                                 ByVal nIndex As Long) As LongPtr

    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" _
                         Alias "SetWindowsHookExA" ( _
                                 ByVal idHook As Long, _
                                 ByVal lpfn As LongPtr, _
                                 ByVal hmod As LongPtr, _
                                 ByVal dwThreadId As Long) As LongPtr

    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
                                 ByVal hHook As LongPtr, _
                                 ByVal nCode As Long, _
                                 ByVal wParam As LongPtr, _
                                 ByRef lParam As Any) As LongPtr

    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
                                 ByVal hHook As LongPtr) As Long

    Private Declare PtrSafe Function PostMessage Lib "user32.dll" _
                         Alias "PostMessageA" ( _
                                 ByVal hwnd As LongPtr, _
                                 ByVal wMsg As Long, _
                                 ByVal wParam As LongPtr, _
                                 ByVal lParam As LongPtr) As Long

    Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                 ByVal xPoint As Long, _
                                 ByVal yPoint As Long) As LongPtr

    Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" ( _
                                 ByRef lpPoint As POINTAPI) As Long
#Else
    Private Declare Function FindWindow Lib "user32" _
                         Alias "FindWindowA" ( _
                                 ByVal lpClassName As String, _
                                 ByVal lpWindowName As String) As Long

    Private Declare Function GetWindowLong Lib "user32.dll" _
                         Alias "GetWindowLongA" ( _
                                 ByVal hwnd As Long, _
                                 ByVal nIndex As Long) As Long

    Private Declare Function SetWindowsHookEx Lib "user32" _
                         Alias "SetWindowsHookExA" ( _
                                 ByVal idHook As Long, _
                                 ByVal lpfn As Long, _
                                 ByVal hmod As Long, _
                                 ByVal dwThreadId As Long) As Long

    Private Declare Function CallNextHookEx Lib "user32" ( _
                                 ByVal hHook As Long, _
                                 ByVal nCode As Long, _
                                 ByVal wParam As Long, _
                                 ByRef lParam As Any) As Long

    Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                                 ByVal hHook As Long) As Long

    Private Declare Function PostMessage Lib "user32.dll" _
                         Alias "PostMessageA" ( _
                                 ByVal hwnd As Long, _
                                 ByVal wMsg As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long

    Private Declare Function WindowFromPoint Lib "user32" ( _
                                 ByVal xPoint As Long, _
                                 ByVal yPoint As Long) As Long

    Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                 ByRef lpPoint As POINTAPI) As Long
#End If
   
Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)
   
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28
Private Const WM_LBUTTONDOWN As Long = &H201
   
Private mLngMouseHook As LongPtr
Private mListBoxHwnd As LongPtr
Private mbHook As Boolean
   
Sub HookListBoxScroll()
Dim lngAppInst As LongPtr
Dim hwndUnderCursor As LongPtr
Dim tPT As POINTAPI
        GetCursorPos tPT
        hwndUnderCursor = WindowFromPoint(tPT.x, tPT.y)
        If mListBoxHwnd <> hwndUnderCursor Then
             UnhookListBoxScroll
             mListBoxHwnd = hwndUnderCursor
                lngAppInst = GetWindowLongPtr(mListBoxHwnd, GWL_HINSTANCE)
                PostMessage mListBoxHwnd, WM_LBUTTONDOWN, 0, 0
             If Not mbHook Then
                     mLngMouseHook = SetWindowsHookEx( _
                                                     WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
                     mbHook = mLngMouseHook <> 0
             End If
     End If
End Sub
   
Sub UnhookListBoxScroll()
     If mbHook Then
             UnhookWindowsHookEx mLngMouseHook
             mLngMouseHook = 0
             mListBoxHwnd = 0
             mbHook = False
     End If
End Sub
   
Private Function MouseProc( _
             ByVal nCode As Long, ByVal wParam As LongPtr, _
             ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
        On Error GoTo errH 'Resume Next
        If (nCode = HC_ACTION) Then
             If WindowFromPoint(lParam.pt.x, lParam.pt.y) = mListBoxHwnd Then
                     If wParam = WM_MOUSEWHEEL Then
                             Dim delta As Long
                             delta = (lParam.mouseData And &HFFFF0000) \ &H10000
                             MouseProc = True
                             If delta > 0 Then
                                     PostMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
                                     PostMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
                             ElseIf delta < 0 Then
                                     PostMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
                                     PostMessage mListBoxHwnd, WM_KEYUP, VK_DOWN, 0
                             End If
                             Exit Function
                     End If
             Else
                     UnhookListBoxScroll
             End If
     End If
        MouseProc = CallNextHookEx( _
                             mLngMouseHook, nCode, wParam, ByVal lParam)
     Exit Function
errH:
        UnhookListBoxScroll
End Function

