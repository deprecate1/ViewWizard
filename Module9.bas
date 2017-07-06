Attribute VB_Name = "Module9"
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public ConfigFile As String

'得到设置信息
Public Sub GetSettingsInfo()
NormalSize = GetPrivateProfileInt("Settings", "NormalSize", 1, ConfigFile)
AlwaysOnTop = GetPrivateProfileInt("Settings", "AlwaysOnTop", 1, ConfigFile)
EnableWin = GetPrivateProfileInt("Settings", "EnableWin", 0, ConfigFile)
ShowHideWin = GetPrivateProfileInt("Settings", "ShowHideWin", 0, ConfigFile)
OpenSound = GetPrivateProfileInt("Settings", "OpenSound", 1, ConfigFile)
NotFindHideWin = GetPrivateProfileInt("Settings", "NotFindHideWin", 1, ConfigFile)
AutoFindWin = GetPrivateProfileInt("Settings", "AutoFindWin", 0, ConfigFile)
OpenHideMode = GetPrivateProfileInt("Settings", "OpenHideMode", 0, ConfigFile)
OnTopDraw = GetPrivateProfileInt("Settings", "OnTopDraw", 0, ConfigFile)
DrawType = GetPrivateProfileInt("Settings", "DrawType", 1, ConfigFile)
End Sub

'保存设置信息
Public Sub SaveSettingsInfo()
WritePrivateProfileString "Settings", "NormalSize", CStr(IIf(NormalSize, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "AlwaysOnTop", CStr(IIf(AlwaysOnTop, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "EnableWin", CStr(IIf(EnableWin, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "ShowHideWin", CStr(IIf(ShowHideWin, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "OpenSound", CStr(IIf(OpenSound, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "NotFindHideWin", CStr(IIf(NotFindHideWin, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "AutoFindWin", CStr(IIf(AutoFindWin, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "OpenHideMode", CStr(IIf(OpenHideMode, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "OnTopDraw", CStr(IIf(OnTopDraw, 1, 0)), ConfigFile
WritePrivateProfileString "Settings", "DrawType", CStr(DrawType), ConfigFile
End Sub
