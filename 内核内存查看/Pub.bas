Attribute VB_Name = "Module1"
Public Declare Function StrToIntExA Lib "shlwapi" (ByVal hexStr As String, ByVal dwFlags As Long, FAR As Long) As Boolean
Public Declare Function DeviceIoControl Lib "Kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Any, lpOverlapped As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetLastError Lib "Kernel32" () As Long
Public Dev As New 驱动
Public Function DbgPrintArray(lpArray() As Byte, Optional ByVal lpEndIndex As Long = 0) As String
Dim a As String
     For i = 0 To UBound(lpArray)
          a = Hex(CLng(lpArray(i)))
          If lpEndIndex >= i Then DbgPrintArray = DbgPrintArray & " " & IIf(Len(a) = 1, "0" & a, a)
     Next
     DbgPrintArray = Trim(DbgPrintArray)
End Function

Public Sub ReadMem(lpAddrAndSize() As Long, lpOutBuffer() As Byte, ByVal lpOutSize As Long)
Dim hDevice As Long
Dim BytRet As Long
    hDevice = Dev.连接设备驱动("\\.\HanfSys_Disastrously")
    DeviceIoControl hDevice, _
    Dev.驱动IO控制码(&H22, &H800, METHOD_缓冲方式, FILE_基本权限), lpAddrAndSize(0), 8, lpOutBuffer(0), lpOutSize, BytRet, ByVal 0
    CloseHandle hDevice
End Sub

Public Sub WriteMem(lpAddressAndSizeAndMemData() As Byte, ByVal BytSize As Long, OutBuffer As Long)
Dim hDevice As Long
Dim BytRet As Long
    hDevice = Dev.连接设备驱动("\\.\HanfSys_Disastrously")
    DeviceIoControl hDevice, _
    Dev.驱动IO控制码(&H22, &H801, METHOD_缓冲方式, FILE_基本权限), lpAddressAndSizeAndMemData(0), BytSize, OutBuffer, 4, BytRet, ByVal 0
    CloseHandle hDevice
End Sub

