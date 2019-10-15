VERSION 5.00
Begin VB.Form Dialog1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "汇编"
   ClientHeight    =   1245
   ClientLeft      =   2760
   ClientTop       =   3690
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox Check1 
      Caption         =   "NOP 填充"
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   765
      Value           =   1  'Checked
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   270
      Width           =   4245
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   330
      Left            =   3375
      TabIndex        =   1
      Top             =   735
      Width           =   1080
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      Top             =   735
      Width           =   1080
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
Unload Dialog1
End Sub

Private Sub Form_Load()
Check1.Value = CLng(GetSetting("Hanf", "Disasm", "NOP", 0))
Text1 = Form1.ListView1.SelectedItem.ListSubItems(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "Hanf", "Disasm", "NOP", Check1.Value
End Sub

Private Sub OKButton_Click()
Dim Hexint As Long
StrToIntExA Form1.ListView1.SelectedItem.Text, 1, Hexint
AssembleAddr Hexint

End Sub
