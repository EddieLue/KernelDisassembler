VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Disassembly"
   ClientHeight    =   1665
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   495
      TabIndex        =   1
      Text            =   "0x200"
      Top             =   765
      Width           =   1995
   End
   Begin VB.CommandButton OKButton 
      Appearance      =   0  'Flat
      Caption         =   "确定"
      Height          =   330
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   495
      TabIndex        =   0
      Text            =   "0x00000000"
      Top             =   225
      Width           =   1995
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   330
      Left            =   1455
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "长度:"
      Height          =   285
      Left            =   45
      TabIndex        =   5
      Top             =   765
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "地址:"
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   225
      Width           =   465
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldAddress As String
Dim OldSize As String

Private Sub CancelButton_Click()
Unload Dialog
End Sub

Private Sub Form_Load()
OldAddress = IIf(OldAddress = "", "0x00000000", OldAddress)
OldSize = IIf(OldSize = "", "0x200", OldSize)
Text1 = OldAddress
Text2 = OldSize
End Sub

Private Sub Form_Unload(Cancel As Integer)
OldSize = Text2
OldAddress = Text1
End Sub


Private Sub OKButton_Click()

Dim DisassemblyAddr As Long
Dim Sizes As Long

StrToIntExA IIf(Left(Text2, 2) <> "0x", "0x" & Text2, Text2), 1, Sizes
If Sizes = 0 Then MsgBox "请输入有效长度！", 48, "警告": Exit Sub


StrToIntExA IIf(Left(Text1, 2) <> "0x", "0x" & Text1, Text1), 1, DisassemblyAddr
If DisassemblyAddr = 0 Or (DisassemblyAddr > 0 And DisassemblyAddr < MEM_4MB_PAGES) Then MsgBox "请输入有效内核地址！", 48, "警告": Exit Sub
StartDisassembly DisassemblyAddr, Sizes
Unload Dialog
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
OKButton_Click
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
OKButton_Click
End If
End Sub
