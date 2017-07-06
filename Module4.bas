Attribute VB_Name = "Module4"
Option Explicit

'提升权限所用到的函数
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetProcessHandleCount Lib "kernel32" (ByVal ProcessHandle As Long, nCount As Long) As Long
Public Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String

Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Public Const THREAD_TERMINATE = (&H1)

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const TOKEN_ASSIGN_PRIMARY = &H1
Public Const TOKEN_DUPLICATE = (&H2)
Public Const TOKEN_IMPERSONATE = (&H4)
Public Const TOKEN_QUERY = (&H8)
Public Const TOKEN_QUERY_SOURCE = (&H10)
Public Const TOKEN_ADJUST_PRIVILEGES = (&H20)
Public Const TOKEN_ADJUST_GROUPS = (&H40)
Public Const TOKEN_ADJUST_DEFAULT = (&H80)
Public Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or _
                          TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or _
                          TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const ANYSIZE_ARRAY = 1
Public Type LUID
    lowpart As Long
    highpart As Long
End Type
Public Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type


'这个函数用于在WIN2000系统中，本进程提升权限
Sub SetAdjustPrivileges()
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lp As Long
    hdlProcessHandle = GetCurrentProcess()
    lp = OpenProcessToken(hdlProcessHandle, TOKEN_ALL_ACCESS, hdlTokenHandle)
    lp = LookupPrivilegeValue("", "SeDebugPrivilege", tmpLuid)
    tkp.PrivilegeCount = 1
    tkp.Privileges(0).pLuid = tmpLuid
    tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    lp = AdjustTokenPrivileges(hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded)
End Sub

