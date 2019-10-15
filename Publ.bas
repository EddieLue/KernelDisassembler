Attribute VB_Name = "Publ"
Public Declare Function StrToIntExA Lib "shlwapi" (ByVal hexStr As String, ByVal dwFlags As Long, FAR As Long) As Boolean
Public Declare Function SafeArrayGetDim Lib "oleaut32" (lpArray() As Any) As Boolean
Public Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Any, lpOverlapped As Long) As Long
Public Declare Function VB_SetOptions Lib "Disasm.dll" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Boolean
Public Declare Function Disasm Lib "Disasm.dll" (ByVal Address As Long, ByVal p2 As Long, ByVal Address As Long, p4 As Any, p5 As Any) As Long
Public Declare Function Assemble Lib "Disasm.dll" (ByVal strCode As String, ByVal dwStartAddr As Long, _
lpRetByt As Any, ByVal p4 As Long, ByVal p5 As Long, lp6 As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Const MEM_4MB_PAGES = &H80000000
Public lOldAddress As Long
Public lOldSize As Long
Public Dev As New 驱动
Public Sub StartDisassembly(ByVal DisassemblyAddr As Long, ByVal Sizes As Long)
Dim OutBuffer() As Byte
Dim OutBuffer1() As Byte
Dim InBuffer(1) As Long

InBuffer(0) = DisassemblyAddr
InBuffer(1) = Sizes
ReDim OutBuffer1(Sizes - 1 + 4)
ReDim OutBuffer(Sizes)
Dim sjSize As Long

ReadMem InBuffer(), OutBuffer1(), Sizes + 4
CopyMemory sjSize, OutBuffer1(0), 4
If sjSize <= 0 Then
MsgBox "反汇编失败!", 48, "提示": Exit Sub
Else
If sjSize < Sizes Then MsgBox "部分内存不可读,只能反汇编部分。", 48, "提示"
End If
ReDim OutBuffer(sjSize - 1)
CopyMemory OutBuffer(0), OutBuffer1(4), sjSize
'Disasm
Dim DisasmRetCode(&H338) As Byte
Dim Prarm(44) As Byte
Dim RetMemSize As Long
Dim MemSize As Long
Dim addr As Long
Dim disasmstrAddr As Long
Dim TempHexInfo(44) As Byte
Dim TempDisasmInfo(99) As Byte
VB_SetOptions 0, 0, 0, 1
disasmstrAddr = DisassemblyAddr
addr = VarPtr(OutBuffer(0))
MemSize = 0
Form1.ListView1.ListItems.Clear
Do

RetMemSize = Disasm(addr + MemSize, 1024, disasmstrAddr + MemSize, DisasmRetCode(0), Prarm(0))

If RetMemSize >= 0 Then
     CopyMemory TempHexInfo(0), DisasmRetCode(4), 45
     CopyMemory TempDisasmInfo(0), DisasmRetCode(260), 100
     Form1.AddListItems "0x" & Hex(disasmstrAddr + MemSize), StrConv(TempHexInfo, vbUnicode), StrConv(TempDisasmInfo, vbUnicode)
     MemSize = MemSize + RetMemSize
End If

If MemSize >= UBound(OutBuffer) + 1 Then
Exit Do
End If
Loop
lOldSize = UBound(OutBuffer)
lOldAddress = disasmstrAddr
'HEX Dump
Form1.Text1.Text = ""
Dim strHex As String
Dim DisassemblyXDAddr As Long
Dim i As Long
Dim n As Long
Dim hexIndex As Long
Dim tempHexStr As String
Dim tempHexAsc As String
Dim ascHex As String
Dim tempByte As Long
DisassemblyXDAddr = DisassemblyAddr
Form1.Text1 = ""
     For i = 0 To UBound(OutBuffer) Step 16
     tempHexAsc = ""
     strHex = tempHexAsc
     ascHex = strHex
          For n = 0 To 15
          If UBound(OutBuffer) < (hexIndex + n) Then Exit For
          tempByte = OutBuffer(hexIndex + n)
          tempHexStr = Hex(tempByte)
          strHex = strHex & " " & IIf(Len(tempHexStr) = 1, "0" & tempHexStr, tempHexStr)
     
          If tempByte = 0 Or tempByte = &H9 Or tempByte = &H1 Or tempByte = &HF7 Or tempByte = &HFF Then
          tempHexAsc = "."
          Else
               If tempByte = &HA Then
                    If n > 1 Then
                         If OutBuffer(hexIndex + n - 1) = &HD Then
                         tempByte = Asc(".")
                          tempHexAsc = ChrW(tempByte)
                         Else: tempHexAsc = ChrW(tempByte)
                         End If
                    End If
               Else
               tempHexAsc = ChrW(tempByte)
               End If
          End If
          ascHex = ascHex & "" & IIf(tempHexAsc = "", ".", tempHexAsc)
     
          Next
          If (hexIndex + 16) <= UBound(OutBuffer) Then hexIndex = hexIndex + 16
     
          Form1.Text1 = Form1.Text1 & "0x" & Hex(DisassemblyXDAddr) & " :" & strHex _
                    & String(5, " ") & (ascHex) & vbCrLf
          DisassemblyXDAddr = DisassemblyXDAddr + 16
     Next

Erase OutBuffer1
Erase OutBuffer

End Sub


Public Function AssembleAddr(ByVal WriteAddr As Long)
Dim OutBuffer As Long
Dim InBuffer() As Byte
Dim tempStr(255) As Byte
Dim tempBuffer(255) As Byte
Dim OldtempBuffer(255) As Long
Dim BytSize As Long
Dim BytRet As Long
Dim sOldBytSize As Long
Dim isCheck As Integer
Dim OutCount As Long
Dim i As Long
Dim NextAddr As Long

isCheck = Dialog1.Check1.Value
If isCheck = 1 Then
     sOldBytSize = GetListCodeSize(WriteAddr)
     If sOldBytSize <= 0 Then MsgBox "汇编失败！", 48, "提示": Exit Function
End If
VB_SetOptions 0, 0, 0, 1
BytSize = Assemble(Dialog1.Text1, WriteAddr, tempBuffer(0), 0, 0, tempStr(0))
If BytSize <= 0 Then MsgBox "汇编失败,指令可能存在错误！", 48, "提示": Exit Function

'If Select Check1'add 'Nop'
OutCount = UBound(tempBuffer)
If isCheck = 1 And BytSize < sOldBytSize Then
     For i = 0 To (sOldBytSize - BytSize)
          tempBuffer(BytSize + i) = &H90
     Next
     BytSize = BytSize + (sOldBytSize - BytSize)
End If

ReDim InBuffer(BytSize - 1 + 4 * 2)
CopyMemory InBuffer(0), WriteAddr, 4
CopyMemory InBuffer(4), BytSize, 4
CopyMemory InBuffer(8), tempBuffer(0), BytSize

If (MsgBox("修改内核内存非常危险,请确认修改操作！！" & vbCrLf & vbCrLf & "地址:" & "0x" & Hex(WriteAddr) & vbCrLf _
& "字节:" & DbgPrintArray(tempBuffer(), BytSize - 1) & vbCrLf & "长度:" & BytSize, vbOKCancel, "危险") <> vbOK) Then _
Exit Function

BytSize = UBound(InBuffer) + 1
WriteMem InBuffer(), BytSize, OutBuffer

AssembleAddr = OutBuffer
If OutBuffer < 0 Then MsgBox "汇编失败,可能是该地址不可写！", 48, "提示": Exit Function
If lOldAddress = 0 And lOldSize <= 0 Then Exit Function
Dim tempSelIndex As Long
tempSelIndex = Form1.ListView1.SelectedItem.Index
StartDisassembly lOldAddress, lOldSize
Unload Dialog1
Form1.ListView1.SelectedItem.Selected = False
Form1.ListView1.ListItems.Item(tempSelIndex).Selected = True
Form1.ListView1.ListItems.Item(tempSelIndex).EnsureVisible
End Function

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
hDevice = Dev.连接驱动设备("\\.\HanfSys_Disastrously")
DeviceIoControl hDevice, _
Dev.驱动IO控制码(&H22, &H800, METHOD_缓冲方式, FILE_基本权限), lpAddrAndSize(0), _
8, lpOutBuffer(0), lpOutSize, BytRet, ByVal 0
CloseHandle hDevice
End Sub

Public Sub WriteMem(lpAddressAndSizeAndMemData() As Byte, ByVal BytSize As Long, OutBuffer As Long)
Dim hDevice As Long
Dim BytRet As Long
hDevice = Dev.连接驱动设备("\\.\HanfSys_Disastrously")
DeviceIoControl hDevice, _
Dev.驱动IO控制码(&H22, &H801, METHOD_缓冲方式, FILE_基本权限), lpAddressAndSizeAndMemData(0), BytSize, OutBuffer, 4, BytRet, ByVal 0
CloseHandle hDevice
End Sub
Public Function GetListCodeSize(ByVal CurrentAddr As Long) As Long
Dim NextAddr As Long
     If Form1.ListView1.SelectedItem.Index + 1 <= Form1.ListView1.ListItems.Count Then
     StrToIntExA Form1.ListView1.ListItems.Item(Form1.ListView1.SelectedItem.Index + 1).Text, 1, NextAddr
     GetListCodeSize = NextAddr - CurrentAddr
     Else
     Exit Function
     End If
End Function
