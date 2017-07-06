Attribute VB_Name = "Module8"

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type


Public Function GetProcessNameByPID(ByVal dwPID As Long) As String
Dim pe As PROCESSENTRY32
Dim hSnap As Long
Dim szName As String
Dim nRet As Long

hSnap = CreateToolhelp32Snapshot(&H2, 0)
pe.dwSize = Len(pe)
nRet = Process32First(hSnap, pe)
Do While nRet <> 0
    If dwPID = pe.th32ProcessID Then
        szName = CheckStr(pe.szExeFile)
        Exit Do
    End If
    nRet = Process32Next(hSnap, pe)
Loop
CloseHandle hSnap
GetProcessNameByPID = szName
End Function

'È¥³ý¶àÓà×Ö·û´®
Private Function CheckStr(ByVal sText As String) As String
n = InStr(sText, Chr$(0))
If n <> 0 Then
    CheckStr = Left$(sText, n - 1)
End If
End Function
