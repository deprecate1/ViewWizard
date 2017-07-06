Attribute VB_Name = "Module6"
Option Explicit

Private Type SYSTEM_MODULE_INFORMATION
    dwReserved(1) As Long
    dwBase As Long
    dwSize As Long
    dwFlags As Long
    Index As Integer
    Unknown As Integer
    LoadCount As Integer
    ModuleNameOffset As Integer
    ImageName As String * 256
End Type

Private Type MODULE_INFO
    dwBase As String
    szModulePath As String
End Type

Private Type MODULES
    dwNumberOfModules As Long
    ModuleInformation As SYSTEM_MODULE_INFORMATION
End Type

Private Declare Function NtQuerySystemInformation Lib "NTDLL.DLL" ( _
            ByVal SystemInformationClass As Long, _
            ByVal pSystemInformation As Long, _
            ByVal SystemInformationLength As Long, _
            ByRef ReturnLength As Long) As Long
            
Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByVal Address As Long, ByVal dwSize As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            ByVal pDst As Long, _
            ByVal pSrc As Long, _
            ByVal ByteLen As Long)

Private Const SystemModuleInformation = 11
Private Const PAGE_READWRITE = &H4
Private Const MEM_RELEASE = &H8000
Private Const MEM_COMMIT = &H1000

Public KernelModules() As MODULE_INFO

Public Function EnumKernelModules() As Long
Dim Ret As Long
Dim Buffer As Long
Dim ModulesInfo As MODULES
Dim i As Long
Dim k As Long
Erase KernelModules
NtQuerySystemInformation SystemModuleInformation, 0, 0, Ret
Buffer = VirtualAlloc(0, Ret * 2, MEM_COMMIT, PAGE_READWRITE)
NtQuerySystemInformation SystemModuleInformation, Buffer, Ret * 2, Ret
CopyMemory ByVal VarPtr(ModulesInfo), ByVal Buffer, LenB(ModulesInfo)
i = ModulesInfo.dwNumberOfModules
While (i > 1)
    i = i - 1
    Buffer = Buffer + 71 * 4
    CopyMemory ByVal VarPtr(ModulesInfo), ByVal Buffer, LenB(ModulesInfo)
    k = k + 1
    ReDim Preserve KernelModules(k)
    KernelModules(k).dwBase = ModulesInfo.ModuleInformation.dwBase
    KernelModules(k).szModulePath = CheckPath(CheckStr(StrConv(ModulesInfo.ModuleInformation.ImageName, vbUnicode)))
    If Fe(KernelModules(k).szModulePath) = False Then
       If Fe(GetFullPath(KernelModules(k).szModulePath)) = True Then
          KernelModules(k).szModulePath = GetFullPath(KernelModules(k).szModulePath)
       End If
    End If
Wend
EnumKernelModules = k
End Function

Private Function GetFullPath(ByVal szPath As String) As String
Dim FullPath  As String
FullPath = GetSysDir & "\drivers\" & szPath
If Fe(FullPath) = True Then
   GetFullPath = FullPath
   Exit Function
End If
FullPath = GetSysDir & "\" & szPath
If Fe(FullPath) = True Then
   GetFullPath = FullPath
   Exit Function
End If
End Function

Public Function GetKernelModulePath(ByVal ModuleBase As Long) As String
Dim i As Long
Dim nRet As Long
nRet = EnumKernelModules
If nRet > 0 Then
   For i = 1 To UBound(KernelModules)
      If KernelModules(i).dwBase = ModuleBase Then
         GetKernelModulePath = KernelModules(i).szModulePath
         Exit For
      End If
   Next i
End If
End Function

