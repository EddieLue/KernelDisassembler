VERSION 5.00
Begin VB.Form Dialog2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�����Ʊ༭"
   ClientHeight    =   1215
   ClientLeft      =   2760
   ClientTop       =   3690
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   330
      Left            =   2115
      TabIndex        =   2
      Top             =   735
      Width           =   1080
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   330
      Left            =   3330
      TabIndex        =   1
      Top             =   735
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   4245
   End
End
Attribute VB_Name = "Dialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Dialog2
End Sub

Private Sub Form_Load()
Dim OutBuffer() As Byte
Dim InBuffer(1) As Long
Dim CurAddr As Long
Dim nCodeSize As Long
Dim OutBytSize As Long
Dim OutMembyt() As Byte

StrToIntExA Form1.ListView1.SelectedItem.Text, 1, CurAddr
nCodeSize = GetListCodeSize(CurAddr)
InBuffer(0) = CurAddr
InBuffer(1) = nCodeSize
ReDim OutBuffer(nCodeSize + 4 - 1)

ReadMem InBuffer(), OutBuffer(), nCodeSize + 4

CopyMemory OutBytSize, OutBuffer(0), 4
Debug.Print OutBytSize; nCodeSize
If OutBytSize <> nCodeSize Then Exit Sub

If nCodeSize <= 0 Then Exit Sub
ReDim OutMembyt(OutBytSize - 1)

CopyMemory OutMembyt(0), OutBuffer(4), OutBytSize

Dim tempHexStr As String
Dim tempHex As String
Dim i As Long
For i = 0 To UBound(OutMembyt)
     tempHex = Hex(OutMembyt(i))
     tempHexStr = tempHexStr & " " & IIf(Len(tempHex) = 1, "0" & tempHex, tempHex)
Next
Text1 = Trim(tempHexStr)

End Sub

Private Sub OKButton_Click()
Dim tempStrAry() As String
Dim InBuffer(255) As Byte
Dim WriteAddr As Long
Dim BytSize As Long
Dim tempBuffer() As Byte
Dim tempIntex As Long
Dim OutBuffer As Long
Dim i As Long

StrToIntExA Form1.ListView1.SelectedItem.Text, 1, WriteAddr
If WriteAddr = 0 Then Exit Sub

tempStrAry = Split(Trim(Text1), " ", , vbTextCompare)
If SafeArrayGetDim(tempStrAry()) = False Then Exit Sub
If UBound(tempStrAry) < 0 Then Exit Sub
ReDim tempBuffer(UBound(tempStrAry))
For i = 0 To UBound(tempBuffer)
     StrToIntExA "0x" & tempStrAry(i), 1, tempIntex
     If tempIntex > 256 Then MsgBox "д��ʧ��,��������������", 48, "����": Exit Sub
     tempBuffer(i) = CByte(tempIntex)
Next
BytSize = UBound(tempBuffer) + 1
CopyMemory InBuffer(0), WriteAddr, 4
CopyMemory InBuffer(4), BytSize, 4
CopyMemory InBuffer(8), tempBuffer(0), BytSize

If (MsgBox("�޸��ں��ڴ�ǳ�Σ��,��ȷ���޸Ĳ�������" & vbCrLf & vbCrLf & "��ַ:" & "0x" & Hex(WriteAddr) & vbCrLf _
& "�ֽ�:" & DbgPrintArray(tempBuffer(), BytSize - 1) & vbCrLf & "����:" & BytSize, vbOKCancel, "Σ��") <> vbOK) Then _
Exit Sub


WriteMem InBuffer(), UBound(InBuffer) + 1, OutBuffer
If OutBuffer < 0 Then MsgBox "д��ʧ�ܣ�", 48, "��ʾ": Exit Sub
Dim tempSelIndex As Long
tempSelIndex = Form1.ListView1.SelectedItem.Index
StartDisassembly lOldAddress, lOldSize
Unload Dialog1
Form1.ListView1.SelectedItem.Selected = False
Form1.ListView1.ListItems.Item(tempSelIndex).Selected = True
Form1.ListView1.ListItems.Item(tempSelIndex).EnsureVisible
End Sub

