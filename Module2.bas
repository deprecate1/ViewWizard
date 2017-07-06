Attribute VB_Name = "Module2"
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (lpszName As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Const SND_ASYNC = &H1
Private Const SND_NOWAIT = &H2000
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2
Private Const SND_FILENAME = &H20000
Private Const SND_RESOURCE = &H40004


Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASE = &H4
Public Const RDW_INVALIDATE = &H1
Public Const PS_INSIDEFRAME = 6
Public Const NULL_BRUSH = 5

Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000

' Public constants for SetWindowPos API declaration
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Type POINTAPI
    X As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum FIND_TYPE
        NON_USE = 0
        REAL_LINE = 1
        BLACK_BLOCK = 2
        FOCUS_RECT = 3
End Enum


Public sndFile As String

'高亮显示
Public Sub DrawFrame(ByVal hWnd As Long, ByVal fTopPaint As Boolean, ByVal FindStyle As FIND_TYPE)

                    If FindStyle = NON_USE Then Exit Sub
                      
                    Dim hdc     As Long, xRect       As RECT
                    Dim Pen     As Long, Brush       As Long
                    Dim OldMode     As Long, OldPen       As Long, OldBrush       As Long
                                
                    Call GetWindowRect(hWnd, xRect)                                                 '获取窗口矩形区域
                    hdc = GetWindowDC(IIf(fTopPaint, 0, hWnd))                                                             '获取窗口场景
                              
                    Select Case FindStyle
                          Case FIND_TYPE.REAL_LINE
                              Pen = CreatePen(PS_INSIDEFRAME, 3, vbRed)
                              Brush = GetStockObject(NULL_BRUSH)
                              OldMode = SetROP2(hdc, vbInvert)
                          Case FIND_TYPE.BLACK_BLOCK
                              Pen = CreatePen(PS_INSIDEFRAME, 3, vbRed)
                              Brush = GetStockObject(BLACK_BRUSH)
                              OldMode = SetROP2(hdc, vbInvert)
                    End Select
                    
                    OldPen = SelectObject(hdc, Pen)                                                 '选定画笔
                    OldBrush = SelectObject(hdc, Brush)                                         '选定画刷
                    
                    If fTopPaint = True Then
                        If FindStyle = FOCUS_RECT Then
                            DrawFocusWindow hdc, xRect, True
                        Else
                            Rectangle hdc, xRect.Left, xRect.Top, xRect.Right, xRect.Bottom    '绘制矩形
                        End If
                    Else
                        If FindStyle = FOCUS_RECT Then
                            DrawFocusWindow hdc, xRect, False
                        Else
                            Rectangle hdc, 0, 0, xRect.Right - xRect.Left, xRect.Bottom - xRect.Top
                        End If
                    End If
                      
                    Call SetROP2(hdc, OldMode)                                                         '恢复现场
                    Call SelectObject(hdc, OldPen)
                    Call SelectObject(hdc, OldBrush)
                      
                    DeleteObject Pen                                                                           '删除画笔
                    DeleteObject Brush                                                                       '删除画刷
                    ReleaseDC hWnd, hdc                                                                       '释放场景
  End Sub

'画焦点窗口
Public Sub DrawFocusWindow(ByVal hdc As Long, lpRect As RECT, ByVal fTopPaint As Boolean)
        Dim xRect As RECT
        Dim lRect As RECT
        xRect = lpRect
        If fTopPaint = False Then
            lRect.Left = 0
            lRect.Top = 0
            lRect.Right = xRect.Right - xRect.Left
            lRect.Bottom = xRect.Bottom - xRect.Top
            DrawFocusRect hdc, lRect
            
            lRect.Left = 1
            lRect.Top = 1
            lRect.Right = xRect.Right - xRect.Left - 1
            lRect.Bottom = xRect.Bottom - xRect.Top - 1
            DrawFocusRect hdc, lRect
            
            lRect.Left = 2
            lRect.Top = 2
            lRect.Right = xRect.Right - xRect.Left - 2
            lRect.Bottom = xRect.Bottom - xRect.Top - 2
            DrawFocusRect hdc, lRect
        Else
            lRect = xRect
            DrawFocusRect hdc, lRect
            
            lRect.Left = xRect.Left + 1
            lRect.Top = xRect.Top + 1
            lRect.Right = xRect.Right - 1
            lRect.Bottom = xRect.Bottom - 1
            DrawFocusRect hdc, lRect
            
            lRect.Left = xRect.Left + 2
            lRect.Top = xRect.Top + 2
            lRect.Right = xRect.Right - 2
            lRect.Bottom = xRect.Bottom - 2
            DrawFocusRect hdc, lRect
        End If
End Sub


Public Function GetPWnd(ByVal hWnd As Long) As Long
  Dim tmp As Long
  tmp = hWnd
  Do
    GetPWnd = tmp
    tmp = GetParent(tmp)
  Loop While tmp <> 0
End Function

Public Sub MakeFlat(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    ' Setup window styles
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    ' Set window style
    SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
    
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

'获得临时文件目录
Public Sub CreateSoundFile()
    Dim Temp As String * 260
    Dim tmpFile As String
    Dim sndByte() As Byte
    GetTempPath 260, Temp
    tmpFile = CheckStr(Temp)
    GetTempFileName tmpFile, "snd", 0&, Temp
    tmpFile = CheckStr(Temp)
    sndFile = tmpFile
    'Debug.Print sndFile
    On Error GoTo Err
    sndByte = LoadResData(101, "Custom")
    Open sndFile For Binary As #1
        Put #1, , sndByte
    Close #1
Err:
End Sub

Sub PlaySnd()
On Error Resume Next
Dim dwFlags As Long
Dim sndByte() As Byte
If Fe(sndFile) = True Then
    dwFlags = SND_FILENAME Or SND_ASYNC Or SND_NOWAIT Or SND_NODEFAULT
    sndPlaySound sndFile, 3&
End If
End Sub

