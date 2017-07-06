Attribute VB_Name = "Module1"
Option Base 1

Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal HWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long, ByVal un As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, X As Long, y As Long) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PostMessageByString& Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Public Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long

Public Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
Public Declare Function CopyIcon Lib "user32" (ByVal hcur As Long) As Long
Public Declare Function DestroyCursor Lib "user32" (ByVal hcursor As Long) As Long
Public Declare Function GetCursor Lib "user32" () As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias _
"sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const GWL_ID = (-12)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
Public Const GWL_HINSTANCE = (-6)
Public Const GCW_ATOM = (-32)
Public Const GCL_HMODULE = (-16)
Public Const GW_HWNDNEXT = 2
Public Const GA_PARENT = 1

Public Const WS_POPUP = &H80000000
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1

Public Const WM_GETTEXT = &HD
Public Const WM_SETTEXT = &HC
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_SETFOCUS = &H7
Public Const WM_DESTROY = &H2
Public Const WM_NCDESTROY = &H82
Public Const WM_SYSCOMMAND = &H112
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const EM_GETPASSWORDCHAR = &HD2
Public Const EM_SETPASSWORDCHAR = &HCC

Public Const HTCAPTION = 2

Public Const SC_CLOSE = &HF060&

Public Const OCR_NORMAL = 32512
Public Const CB_ADDSTRING = &H143

Public wCount As Long
Public HwndChild() As Long
Public a As Long, b As Long

Public AlwaysOnTop As Boolean
Public EnableWin As Boolean
Public ShowHideWin As Boolean
Public OpenSound As Boolean
Public NotFindHideWin As Boolean
Public AutoFindWin As Boolean
Public OpenHideMode As Boolean
Public OnTopDraw As Boolean
Public DrawType As FIND_TYPE
Public NormalSize As Boolean

Public hSave As Long

Public Function EnumCallback(ByVal app_hWnd As Long, ByVal param As Long) As Long
On Error Resume Next
    wCount = wCount + 1
    ReDim Preserve HwndChild(wCount)
    HwndChild(wCount) = app_hWnd
    EnumCallback = 1
End Function


'获得弹出式窗口
Function GetParentHwnd(ByVal hWindow As Long) As Long
Dim TmpHwnd As Long
Dim TopWin As Long
TmpHwnd = hWindow
Do
    TopWin = TmpHwnd
    If IsPopupWin(TmpHwnd) = True Then Exit Do
    TmpHwnd = GetParent(TmpHwnd)
Loop While TmpHwnd <> 0
GetParentHwnd = TopWin
End Function

'判断是否为弹出式窗口
Function IsPopupWin(ByVal hWindow As Long) As Boolean
Dim lngStyle As Long
lngStyle = GetWindowLong(hWindow, GWL_STYLE)
If lngStyle And WS_POPUP Then
    IsPopupWin = True
Else
    IsPopupWin = False
End If
End Function

'得到窗口类名
Function GetWinClass(ByVal hWindow As Long) As String
Dim wClass As String * 255
GetClassName hWindow, wClass, 255
GetWinClass = CheckStr(wClass)
End Function

'得到父弹出窗口
Function GetParentW(ByVal hWindow As Long) As Long
Dim pHwnd As Long
pHwnd = GetParent(hWindow)
If pHwnd <> 0 And IsPopupWin(hWindow) = False Then
    GetParentW = pHwnd
Else
    GetParentW = 0
End If
End Function

'得到精确父窗口
Function GetParentEx(ByVal hWin As Long) As Long
Dim hw As Long
Dim hOwner As Long
hw = GetParent(hWin)
If hw <> 0 Then
    GetParentEx = hw
Else
    hw = GetWindowLong(hWin, -8)
    If hw = 0 Then
        GetParentEx = 0
    Else
        hOwner = GetWindow(hWin, 4)
        If hw <> hOwner Then
            GetParentEx = hw
        Else
            GetParentEx = 0
        End If
    End If
End If
End Function

'检查窗口
Function CheckHwnd(ByVal hWindow As Long) As Long
Dim hWnd As Long
Dim pHwnd As Long
Dim SearchHwnd As Long
Dim SearchRect As RECT
Dim wRect As RECT
Dim pt As POINTAPI
Dim i As Long
hWnd = hWindow
pHwnd = GetParentW(hWnd)
If pHwnd = 0 Then pHwnd = hWnd
    wCount = 0
    Erase HwndChild
    EnumChildWindows pHwnd, AddressOf EnumCallback, 0&
    'Debug.Print GetParentHwnd(hWnd), wCount
    If wCount > 0 Then
        GetWindowRect hWnd, wRect
        GetCursorPos pt
        For i = 1 To wCount
            SearchHwnd = HwndChild(i)
            GetWindowRect SearchHwnd, SearchRect
            If CBool(PointInRect(SearchRect, pt)) Then
                If NotFindHideWin = True Then
                        If CBool(IsWindowVisible(SearchHwnd)) Then
                                If (SearchRect.Right - SearchRect.Left) * (SearchRect.Bottom - SearchRect.Top) < _
                                   (wRect.Right - wRect.Left) * (wRect.Bottom - wRect.Top) Then
                                    hWnd = SearchHwnd
                                    GetWindowRect hWnd, wRect
                                End If
                        End If
                Else
                        If (SearchRect.Right - SearchRect.Left) * (SearchRect.Bottom - SearchRect.Top) < _
                                   (wRect.Right - wRect.Left) * (wRect.Bottom - wRect.Top) Then
                                    hWnd = SearchHwnd
                                    GetWindowRect hWnd, wRect
                        End If
                End If
            End If
        Next i
    End If
    CheckHwnd = hWnd

End Function

'判断点是否在矩形内部
Function PointInRect(lpRect As RECT, pt As POINTAPI) As Boolean
If pt.X >= lpRect.Left And pt.X < lpRect.Right And pt.y >= lpRect.Top And pt.y < lpRect.Bottom Then
    PointInRect = True
Else
    PointInRect = False
End If
End Function

Public Function GetWinhInstance(ByVal hWindow As Long) As Long
GetWinhInstance = GetWindowLong(hWindow, GWL_HINSTANCE)
End Function

Public Function GetWinhModule(ByVal hWindow As Long) As Long
GetWinhModule = GetClassLong(hWindow, GCL_HMODULE)
End Function

Public Function GetFullPath(ByVal sPath As String) As String
Dim tmpStr As String
tmpStr = String$(260, vbNullChar)
lngth = GetFullPathName(sPath, Len(tmpStr), tmpStr, ByVal 0&)
GetFullPath = Left$(tmpStr, lngth)
End Function


