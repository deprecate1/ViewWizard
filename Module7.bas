Attribute VB_Name = "Module7"


Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function NtQueryInformationProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As Long, ByVal ProcessInformation As Long, ByVal ProcessInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long

Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = (&H400)

Public Type PROCESS_BASIC_INFORMATION
        ExitStatus  As Long    'NTSTATUS 接收进程终止状态
        PebBaseAddress  As Long    'PPEB 接收进程环境块地址
        AffinityMask  As Long    'ULONG_PTR 接收进程关联掩码
        BasePriority  As Long    'KPRIORITY 接收进程的优先级类
        UniqueProcessId  As Long    'ULONG_PTR 接收进程ID
        InheritedFromUniqueProcessId  As Long    'ULONG_PTR 接收父进程ID
End Type

'取得父进程ID
Public Function GetParentPid(ByVal PID As Long) As Long
Dim objBasic    As PROCESS_BASIC_INFORMATION
Dim hProcess As Long, dwPId As Long
ProcessBasicInformation = 0
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, PID)
If hProcess <> 0 Then
    ntStatus = NtQueryInformationProcess(hProcess, ProcessBasicInformation, VarPtr(objBasic), Len(objBasic), 0)
    dwPId = objBasic.InheritedFromUniqueProcessId
    If dwPId < 0 Then dwPId = 0
    CloseHandle hProcess
Else
    dwPId = 0
End If
GetParentPid = dwPId
End Function



