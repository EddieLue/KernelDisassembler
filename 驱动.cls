VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "驱动"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByVal lpInBuffer As Long, ByVal nInBufferSize As Long, ByVal lpOutBuffer As Long, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Long) As Long
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As Long, ByVal lpLoadOrderGroup As Long, lpdwTagId As Long, ByVal lpDependencies As Long, ByVal lp As Long, ByVal lpPassword As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (lpPrevWndFunc As Any, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type SERVICE_STATUS
        dwServiceType As Long
        dwCurrentState As Long
        dwControlsAccepted As Long
        dwWin32ExitCode As Long
        dwServiceSpecificExitCode As Long
        dwCheckPoint As Long
        dwWaitHint As Long
End Type
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Enum vMETHOD
    METHOD_缓冲方式 = &H0
    METHOD_直接方式输入 = &H1
    METHOD_直接方式输出 = &H2
    METHOD_未知方式 = &H3
End Enum
Enum vFILEACCESS
    FILE_基本权限 = 0
    FILE_可读权限 = 1
    FILE_可写权限 = 2
End Enum
Dim hService As Long
 Dim hSCManager As Long
Public Function 加载驱动(ByVal 驱动路径 As String, ByVal 驱动服务名 As String) As Long
    hSCManager = OpenSCManager(0, 0, 983103)
    hService = CreateService(hSCManager, 驱动服务名, 驱动服务名 & " Display", 983040 Or 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64 Or 128 Or 256, 1, 3, 0, VarPtr(驱动路径), 0, ByVal 0, 0, 0, 0)

    If (hService <= 0) Then hService = OpenService(hSCManager, 驱动服务名, 983040 Or 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64 Or 128 Or 256)
    StartService hService, 0, 0
    加载驱动 = hService
End Function

Public Function 卸载驱动(ByVal 是否删除服务 As Boolean) As Boolean
Dim Retn As Boolean: Dim stsat As SERVICE_STATUS
    Retn = ControlService(hService, 1, stsat) <> 0
    If 是否删除服务 Then DeleteService (hService)
    CloseServiceHandle (hSCManager)
    CloseServiceHandle (hService)
    hService = 0
    卸载驱动 = Retn
End Function

Public Property Get 驱动服务句柄() As Variant
驱动句柄 = hService
End Property

Public Property Let 驱动服务句柄(ByVal vNewValue As Variant)
hService = vNewValue
End Property

Public Function 与驱动通信(ByVal 驱动控制码 As Long, ByVal 输入指针 As Long, _
ByVal 输入指针长度 As Long, ByVal 输出指针 As Long, ByVal 输出指针长度 As Long) As Long
    通信 = DeviceIoControl(hService, 驱动控制码, ByVal 输入指针, 输入指针长度, 输出指针, 输出指针长度, ByVal 0, ByVal 0)
End Function

Public Function 连接设备驱动(ByVal 驱动服务名 As String, Optional ByVal 操作权限 As Long) As Long '默认=全权=269484160  读写=3222274176
Dim temp As SECURITY_ATTRIBUTES
    连接设备 = CreateFile(VarPtr(驱动服务名), IIf(操作权限 = 0, &H844FE458 Or &H40000000, 操作权限), 0, temp, 3, &H80, 0)   '返回设备驱动服务句柄
End Function

Public Function 驱动IO控制码(设备类型, IO控制码, ByVal 内存方式 As vMETHOD, ByVal 访问权限 As vFILEACCESS) As Long
'设备类型:NT驱动=FILE_DEVICE_UNKNOWN=未知设备=0x22=34
'IO控制码:0x0000-0x7FFF=微软保留(0-32767),0x800-0xFFF=用户自定义(2048-4095)
Dim bB() As Byte
ReDim bB(0 To 35)
a = Array(&H55, &H89, &HE5, &H8B, &H55, &H14, &H8B, &H4D, &H10, &H8B, &H5D, &HC, _
&H8B, &H45, &H8, &HC1, &HE0, &H10, &HC1, &HE2, &HE, &HC1, &HE3, &H2, &H9, &HD0, &H9, _
&HD8, &H9, &HC8, &H89, &HEC, &H5D, &HC2, &H10, &H0)
ReDim bB(0 To UBound(a))
For i = 0 To UBound(a)
bB(i) = CByte(a(i))
Next
驱动IO控制码 = CallWindowProc(bB(0), 设备类型, IO控制码, 内存方式, 访问权限)
ReDim bB(0)
End Function
