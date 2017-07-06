Attribute VB_Name = "Module3"
'Option Explicit


Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_VM_READ = 16
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_SET_INFORMATION = 612
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_SET_SESSIONID = (&H4)
Public Const MAX_PATH As Integer = 260

Public nThread As Long
Public nTemp As Integer

'获得系统system32目录
Public Function GetSysDir() As String
    Dim Temp As String * 256
    Dim X As Integer
    X = GetSystemDirectory(Temp, Len(Temp))
    GetSysDir = Left$(Temp, X)
End Function

'获得Win目录
Public Function GetWinDir() As String
    Dim Temp As String * 256
    Dim X As Integer
    X = GetWindowsDirectory(Temp, Len(Temp))
    GetWinDir = Left$(Temp, X)
End Function

Public Function CheckPath(ByVal PathStr As String) As String
On Error Resume Next
    PathStr = Replace(PathStr, "\??\", "")
    If UCase(Left$(PathStr, 12)) = "\SYSTEMROOT\" Then PathStr = GetWinDir & Mid$(PathStr, 12)
    PathStr = Trim$(PathStr)
    CheckPath = PathStr
End Function

Public Function GetProcessPath(ByVal ProcessId As Long) As String
Dim lngModules(1 To 200) As Long
Dim lngCBSize2 As Long
Dim lngReturn As Long
Dim strModuleName As String
Dim strProcessName As String
Dim hProcess As Long
  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessId)
    If hProcess <> 0 Then
    lngReturn = EnumProcessModules(hProcess, lngModules(1), 200, lngCBSize2)
    strModuleName = Space(MAX_PATH)
    lngReturn = GetModuleFileNameExA(hProcess, lngModules(1), strModuleName, 500)
    strProcessName = Left(strModuleName, lngReturn)
    strProcessName = CheckPath(Trim$(strProcessName))
    GetProcessPath = GetFullPath(strProcessName)
    CloseHandle hProcess
    Else
    GetProcessPath = ""
    End If
End Function

Public Function GetModulePath(ByVal hWnd As Long) As String
Dim PID As Long
Dim sModuleName As String
Dim sPath As String
Dim hModule As Long
Dim hProcess As Long
hModule = GetWinhInstance(hWnd)
GetWindowThreadProcessId hWnd, PID
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
If hProcess <> 0 Then
    sModuleName = Space(260)
    Call GetModuleFileNameExA(hProcess, hModule, sModuleName, 260)
    sPath = CheckPath(CheckStr(sModuleName))
    If Fe(sPath) Then
       GetModulePath = GetFullPath(sPath)
    Else
       GetModulePath = GetFullPath(GetKernelModulePath(hModule))
    End If
    CloseHandle hProcess
End If
End Function

Public Function GetModuleName(ByVal hWnd As Long) As String
GetModuleName = PathToName(GetModulePath(hWnd))
End Function

Public Function GetProcessName(ByVal ProcessId As Long) As String
tmp = GetProcessPath(ProcessId)
If tmp <> "" Then
  GetProcessName = Mid$(tmp, InStrRev(tmp, "\") + 1)
Else
  GetProcessName = tmp
End If
End Function

Function Fe(ByVal szFileName As String) As Boolean
If PathFileExists(szFileName) > 0 Then
   If GetAttr(szFileName) And vbDirectory Then
      Fe = False
   Else
      Fe = True
   End If
Else
   Fe = False
End If
End Function

Public Function PathToName(ByVal sPath As String) As String
n = InStrRev(sPath, "\")
If n <> 0 And n < Len(sPath) Then
   PathToName = Mid$(sPath, n + 1)
End If
End Function

Public Function CheckStr(ByVal sText As String) As String
n = InStr(sText, Chr$(0))
If n <> 0 Then
   CheckStr = Left$(sText, n - 1)
Else
   CheckStr = sText
End If
End Function

Public Function LButtonDown() As Boolean
If GetAsyncKeyState(1) = -32767 Or GetAsyncKeyState(1) = -32768 Then
   LButtonDown = True
Else
   LButtonDown = False
End If
End Function


