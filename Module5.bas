Attribute VB_Name = "Module5"
'窗口一般样式(16)
Private Const WS_BORDER = &H800000
Private Const WS_CHILD = &H40000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_HSCROLL = &H100000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000

'窗口扩展样式(20)
Private Const WS_EX_WINDOWEDGE = &H100&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TOOLWINDOW = &H80&
Private Const WS_EX_STATICEDGE = &H20000
Private Const WS_EX_RTLREADING = &H2000&
Private Const WS_EX_RIGHT = &H1000&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_NOINHERITLAYOUT = &H100000
Private Const WS_EX_NOACTIVATE = &H8000000
Private Const WS_EX_MDICHILD = &H40&
Private Const WS_EX_LEFTSCROLLBAR = &H4000&
Private Const WS_EX_LAYOUTRTL = &H400000
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_CONTROLPARENT = &H10000
Private Const WS_EX_CONTEXTHELP = &H400&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_ACCEPTFILES = &H10&

'得到窗口样式
Public Function GetWindowStyle(ByVal lhWnd As Long) As String
    Dim lStyle As Long
    Dim TmpVal As String
    Dim winStyle As String
    ' Get window styles
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    
    ' Get window styles
    If (lStyle And WS_BORDER) = WS_BORDER Then winStyle = winStyle & "WS_BORDER|"
    If (lStyle And WS_CHILD) = WS_CHILD Then winStyle = winStyle & "WS_CHILD|"
    If (lStyle And WS_CLIPCHILDREN) = WS_CLIPCHILDREN Then winStyle = winStyle & "WS_CLIPCHILDREN|"
    If (lStyle And WS_CLIPSIBLINGS) = WS_CLIPSIBLINGS Then winStyle = winStyle & "WS_CLIPSIBLINGS|"
    If (lStyle And WS_DISABLED) = WS_DISABLED Then winStyle = winStyle & "WS_DISABLED|"
    If (lStyle And WS_DLGFRAME) = WS_DLGFRAME Then winStyle = winStyle & "WS_DLGFRAME|"
    If (lStyle And WS_HSCROLL) = WS_HSCROLL Then winStyle = winStyle & "WS_HSCROLL|"
    If (lStyle And WS_MAXIMIZE) = WS_MAXIMIZE Then winStyle = winStyle & "WS_MAXIMIZE|"
    If (lStyle And WS_MINIMIZE) = WS_MINIMIZE Then winStyle = winStyle & "WS_MINIMIZE|"
    If (lStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Then winStyle = winStyle & "WS_MAXIMIZEBOX|"
    If (lStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then winStyle = winStyle & "WS_MINIMIZEBOX|"
    If (lStyle And WS_SYSMENU) = WS_SYSMENU Then winStyle = winStyle & "WS_SYSMENU|"
    If (lStyle And WS_POPUP) = WS_POPUP Then winStyle = winStyle & "WS_POPUP|"
    If (lStyle And WS_THICKFRAME) = WS_THICKFRAME Then winStyle = winStyle & "WS_THICKFRAME|"
    If (lStyle And WS_VISIBLE) = WS_VISIBLE Then winStyle = winStyle & "WS_VISIBLE|"
    If (lStyle And WS_VSCROLL) = WS_VSCROLL Then winStyle = winStyle & "WS_VSCROLL|"
    
    TmpVal = Trim$(winStyle)
    If Right$(TmpVal, 1) = "|" Then TmpVal = Left$(TmpVal, Len(TmpVal) - 1)
    GetWindowStyle = TmpVal
End Function

'得到窗口扩展样式
Public Function GetWindowExStyle(ByVal lhWnd As Long) As String
    Dim lStyle As Long
    Dim TmpVal As String
    Dim winExStyle As String
    ' Get window styles
    lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    
    ' Get window styles
    If (lStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then winExStyle = winExStyle & "WS_EX_WINDOWEDGE|"
    If (lStyle And WS_EX_TRANSPARENT) = WS_EX_TRANSPARENT Then winExStyle = winExStyle & "WS_EX_TRANSPARENT|"
    If (lStyle And WS_EX_TOPMOST) = WS_EX_TOPMOST Then winExStyle = winExStyle & "WS_EX_TOPMOST|"
    If (lStyle And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW Then winExStyle = winExStyle & "WS_EX_TOOLWINDOW|"
    If (lStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then winExStyle = winExStyle & "WS_EX_STATICEDGE|"
    If (lStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then winExStyle = winExStyle & "WS_EX_RTLREADING|"
    If (lStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then winExStyle = winExStyle & "WS_EX_RIGHT|"
    If (lStyle And WS_EX_NOPARENTNOTIFY) = WS_EX_NOPARENTNOTIFY Then winExStyle = winExStyle & "WS_EX_NOPARENTNOTIFY|"
    If (lStyle And WS_EX_NOINHERITLAYOUT) = WS_EX_NOINHERITLAYOUT Then winExStyle = winExStyle & "WS_EX_NOINHERITLAYOUT|"
    If (lStyle And WS_EX_NOACTIVATE) = WS_EX_NOACTIVATE Then winExStyle = winExStyle & "WS_EX_NOACTIVATE|"
    If (lStyle And WS_EX_MDICHILD) = WS_EX_MDICHILD Then winExStyle = winExStyle & "WS_EX_MDICHILD|"
    If (lStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then winExStyle = winExStyle & "WS_EX_LEFTSCROLLBAR|"
    If (lStyle And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then winExStyle = winExStyle & "WS_EX_LAYOUTRTL|"
    If (lStyle And WS_EX_LAYERED) = WS_EX_LAYERED Then winExStyle = winExStyle & "WS_EX_LAYERED|"
    If (lStyle And WS_EX_DLGMODALFRAME) = WS_EX_DLGMODALFRAME Then winExStyle = winExStyle & "WS_EX_DLGMODALFRAME|"
    If (lStyle And WS_EX_CONTROLPARENT) = WS_EX_CONTROLPARENT Then winExStyle = winExStyle & "WS_EX_CONTROLPARENT|"
    If (lStyle And WS_EX_CONTEXTHELP) = WS_EX_CONTEXTHELP Then winExStyle = winExStyle & "WS_EX_CONTEXTHELP|"
    If (lStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then winExStyle = winExStyle & "WS_EX_CLIENTEDGE|"
    If (lStyle And WS_EX_APPWINDOW) = WS_EX_APPWINDOW Then winExStyle = winExStyle & "WS_EX_APPWINDOW|"
    If (lStyle And WS_EX_ACCEPTFILES) = WS_EX_ACCEPTFILES Then winExStyle = winExStyle & "WS_EX_ACCEPTFILES|"
    
    TmpVal = Trim$(winExStyle)
    If Right$(TmpVal, 1) = "|" Then TmpVal = Left$(TmpVal, Len(TmpVal) - 1)
    GetWindowExStyle = TmpVal
End Function
