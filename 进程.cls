VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "进程"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Type LARGE_INTEGER
        lowpart As Long
        highpart As Long
End Type
Private Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        lowpart  As Long
        highpart  As Long
        Attributes  As Long
End Type
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, _
ByVal lpCommandLine As String, lpProcessAttributes As Long, lpThreadAttributes As Long, _
ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As Long, _
lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type
Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function ZwSuspendProcess Lib "ntdll.dll" (ByVal hProcess As Long) As Long
Private Declare Function ZwResumeProcess Lib "ntdll.dll" (ByVal hProcess As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (lpPrevWndFunc As Any, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As LPPROCESSENTRY32) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As LPPROCESSENTRY32) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Type LPPROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    ProcessName(255) As Byte
End Type
Private Const CP_ACP As Long = 0
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, _
 lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, _
lpUsedDefaultChar As Long) As Long

Private Const MEM_COMMIT = &H1000
Public Function 打开进程(ByVal ProcessID As Long) As Long
    打开进程 = OpenProcess(2035711, 0, ProcessID)
End Function

Public Function 关闭句柄(ByVal 句柄 As Long) As Boolean
    关闭句柄 = CloseHandle(句柄) <> 0
End Function

Public Function 终止(ByVal 进程句柄 As Long) As Long
Dim c As Long
    TerminateProcess 进程句柄, c
    终止 = c
End Function
Public Sub 提升权限(Optional ByVal 进程句柄 As Long, Optional 权限 As String)
Dim wh, lpjb As Long
Dim tmpLuid As LARGE_INTEGER
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    wh = IIf(进程句柄 <= 0, -1, 进程句柄)
    OpenProcessToken wh, 40, lpjb
    LookupPrivilegeValue "", IIf(权限 = "", "SeDebugPrivilege", 权限), tmpLuid
    tkp.PrivilegeCount = 1
    tkp.lowpart = tmpLuid.lowpart
    tkp.highpart = tmpLuid.highpart
    tkp.Attributes = 2
    AdjustTokenPrivileges lpjb, 0, tkp, 28, tkpNewButIgnored, 0
End Sub

Public Function 创建(ByVal 程序文件名 As String, ByVal 命令行 As String, 进程句柄 As Long, _
线程句柄 As Long, Optional 进程ID As Long, Optional 线程ID As Long) As Long
Dim a As STARTUPINFO
Dim b As PROCESS_INFORMATION
    CreateProcess 0, 程序文件名 & " " & 命令行, ByVal 0, ByVal 0, 0, 0, ByVal 0, 0, a, b
    进程句柄 = b.hProcess
    线程句柄 = b.hThread
    进程ID = b.dwProcessId
    线程ID = b.dwThreadId
End Function

Public Function 取自ID() As Long
    取自ID = GetCurrentProcessId()
End Function

Public Function 暂停(ByVal 进程句柄 As Long, Optional ByVal 恢复 As Boolean) As Boolean
If 恢复 Then
    暂停 = ZwResumeProcess(进程句柄) = 0
    Else
    暂停 = ZwSuspendProcess(进程句柄) = 0
    End If
End Function '

Public Function 调用远程入口(ByVal 进程句柄 As Long, ByVal 入口地址 As Long) As Long
Dim a As Long
Dim lretn As Long
    a = CreateRemoteThread(进程句柄, ByVal 0, 0, ByVal 入口地址, ByVal 0, 0, 0)
    WaitForSingleObject a, 255555
    GetExitCodeThread a, lretn
    关闭句柄 (a)
    调用远程入口 = lretn
End Function

Public Function 调用本地入口(ByVal 入口地址 As Long)
    调用本地入口 = CallWindowProc(入口地址, 0, 0, 0, 0)
End Function

Public Function 调用函数(ByVal 进程句柄 As Long, ByVal 地址 As Long, ByRef 函数参数() As Long, _
Optional ByVal 不返回 As Boolean, Optional ByVal 清除调用代码 As Boolean) As Long
Dim addr As Long
Dim codeAry() As Byte
Dim i As Long
ReDim codeAry(4)
codeAry(0) = &HB8
If SafeArrayGetDim(codeAry) <= 0 Or 地址 <= 0 Then Exit Function
    CopyMemory codeAry(1), 地址, 4
        If SafeArrayGetDim(函数参数) > 0 Then
        ReDim Preserve codeAry(4 + ((UBound(函数参数) + 1) * 5))
        For i = UBound(函数参数) To 0 Step -1
            codeAry(4 + (i * 5) + 1) = &H68
            CopyMemory codeAry(4 + ((UBound(函数参数) - i) * 5) + 2), 函数参数(i), 4
        Next
        End If
    ReDim Preserve codeAry(UBound(codeAry) + 2)
    codeAry(UBound(codeAry) - 1) = &HFF: codeAry(UBound(codeAry)) = &HD0
    If 不返回 Then
        ReDim Preserve codeAry(UBound(codeAry) + 2)
        codeAry(UBound(codeAry) - 1) = &H31: codeAry(UBound(codeAry)) = &HC0
    End If
    ReDim Preserve codeAry(UBound(codeAry) + 3)
    codeAry(UBound(codeAry) - 2) = &HC2: codeAry(UBound(codeAry) - 1) = IIf(进程句柄 <= 0, &H10, &H4): codeAry(UBound(codeAry)) = &H0

If 进程句柄 <= 0 Then
    调用函数 = CallWindowProc(codeAry(0), 0, 0, 0, 0)
Else
    Dim hThread As Long
    Dim dwExitCode As Long
    addr = VirtualAllocEx(进程句柄, ByVal 0, UBound(codeAry) + 1, MEM_COMMIT, &H40)
    If addr > 0 Then
        If (WriteProcessMemory(进程句柄, ByVal addr, codeAry(0), UBound(codeAry) + 1, 0) <> 0) Then
            hThread = CreateRemoteThread(进程句柄, ByVal 0, 0, ByVal addr, ByVal 0, 0, ByVal 0)
            If 不返回 = False Then
                WaitForSingleObject hThread, 5000
                GetExitCodeThread hThread, dwExitCode
            End If
            调用函数 = dwExitCode
        End If
    VirtualFreeEx 进程句柄, ByVal 0, UBound(codeAry) + 1, &H8000
    End If
End If
If 清除调用代码 Then Erase codeAry
End Function

Public Function 调用代码(ByVal 进程句柄 As Long, ByRef 代码() As Byte, _
Optional ByVal 不返回 As Boolean, Optional ByVal 清除调用代码 As Boolean) As Long
Dim addr As Long
Dim codeAry() As Byte
ReDim codeAry(UBound(代码))
For i = 0 To UBound(codeAry)
    codeAry(i) = 代码(i)
Next
If 不返回 Then
    ReDim Preserve codeAry(UBound(codeAry) + 2)
    codeAry(UBound(codeAry) - 1) = &H31: codeAry(UBound(codeAry)) = &HC0
End If
ReDim Preserve codeAry(UBound(codeAry) + 3)
codeAry(UBound(codeAry) - 2) = &HC2: codeAry(UBound(codeAry) - 1) = IIf(进程句柄 <= 0, &H10, &H4): codeAry(UBound(codeAry)) = &H0

If 进程句柄 <= 0 Then
    调用代码 = CallWindowProc(codeAry(0), 0, 0, 0, 0)
Else
Dim hThread As Long
Dim dwExitCode As Long
addr = VirtualAllocEx(进程句柄, ByVal 0, UBound(codeAry) + 1, MEM_COMMIT, &H40)
If addr > 0 Then
    If (WriteProcessMemory(进程句柄, ByVal addr, codeAry(0), UBound(codeAry) + 1, 0) <> 0) Then
        hThread = CreateRemoteThread(进程句柄, ByVal 0, 0, ByVal addr, ByVal 0, 0, ByVal 0)
        If 不返回 = False Then
            WaitForSingleObject hThread, 5000
            GetExitCodeThread hThread, dwExitCode
        End If
        调用代码 = dwExitCode
    End If
VirtualFreeEx 进程句柄, ByVal 0, UBound(codeAry) + 1, &H8000
End If
End If
If 清除调用代码 Then Erase codeAry
End Function

Public Function AscArrayToStr(lpByteArray() As Byte) As String
Dim bytAry() As Byte
    ReDim bytAry(UBound(lpByteArray))
For i = 0 To UBound(lpByteArray)
    bytAry(i) = lpByteArray(i)
    If lpByteArray(i) = 0 Then
        ReDim Preserve bytAry(i - 1)
        AscArrayToStr = StrConv(bytAry, vbUnicode)
        Exit For
    End If
Next
End Function

Public Function StrToAscArray(ByVal lpStr As String, lpAscArray() As Byte) As Long
Dim a As Long
Dim strByte() As Byte
Dim b() As Byte
strByte = lpStr
ReDim b(UBound(strByte) / 2)
    a = WideCharToMultiByte(CP_ACP, 0, strByte(0), (UBound(strByte) + 1) / 2, b(0), 0, 0, 0)
ReDim b(a)
    a = WideCharToMultiByte(CP_ACP, 0, strByte(0), -1, b(0), a, 0, 0)
ReDim lpAscArray(UBound(b))
lpAscArray = b
StrToAscArray = UBound(lpAscArray)
End Function

Public Function 枚举PID(PID数组() As Long, 进程名() As String, Optional 指定进程名$) As Long
Dim lppe As LPPROCESSENTRY32
Dim c As Long
Dim i As Long, ProcName() As Byte
Dim s As Long
hSnapshot = CreateToolhelp32Snapshot(2, 0)
lppe.dwSize = 1024
c = Process32First(hSnapshot, lppe)
Dim bProcName() As Byte
While c <> 0
If 指定进程名$ = "" Then
    ReDim Preserve PID数组(s + 1)
    PID数组(s) = lppe.ProcessID
    ReDim Preserve 进程名(s + 1)
    进程名(s) = AscArrayToStr(lppe.ProcessName)
    s = s + 1
Else
    If AscArrayToStr(lppe.ProcessName) = 指定进程名$ Then
        ReDim Preserve PID数组(s)
        PID数组(s) = lppe.ProcessID
        ReDim Preserve 进程名(s + 1)
        进程名(s) = AscArrayToStr(lppe.ProcessName)
        s = s + 1
    End If
End If
    c = Process32Next(hSnapshot, lppe)
Wend
枚举PID = s - 1
End Function
Public Function 进程名取PID(进程名$) As Long
Dim lppe As LPPROCESSENTRY32
Dim c As Long
Dim i As Long, ProcName() As Byte
Dim s As Long
hSnapshot = CreateToolhelp32Snapshot(2, 0)
lppe.dwSize = 1024
c = Process32First(hSnapshot, lppe)
Dim bProcName() As Byte
While c <> 0
ReDim bProcName(UBound(lppe.ProcessName))
For i = 0 To UBound(lppe.ProcessName)
    bProcName(i) = lppe.ProcessName(i)
    If lppe.ProcessName(i) = 0 Then
        ReDim Preserve bProcName(i - 1)
        If StrConv(bProcName, vbUnicode) = 进程名$ Then
            进程名取PID = lppe.ProcessID
        End If
        Exit For
    End If
Next
c = Process32Next(hSnapshot, lppe)
Wend
End Function
