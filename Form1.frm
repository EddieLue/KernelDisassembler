VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "寒风 内核反汇编"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9195
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   503
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "By Leaf Email:958570606@qq.com"
            TextSave        =   "By Leaf Email:958570606@qq.com"
            Object.ToolTipText     =   "作者联系信息"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2940
      Width           =   9150
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2760
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   4868
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   " "
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hex Dump"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Disassembly"
         Object.Width           =   9701
      EndProperty
   End
   Begin VB.Menu m_m 
      Caption         =   "菜单"
      Visible         =   0   'False
      Begin VB.Menu m_Disassembly 
         Caption         =   "反汇编"
      End
      Begin VB.Menu m_1 
         Caption         =   "-"
      End
      Begin VB.Menu m_Goto 
         Caption         =   "跟随"
      End
      Begin VB.Menu m_Copy 
         Caption         =   "复制"
         Begin VB.Menu m_Copy_Addr 
            Caption         =   "地址"
         End
         Begin VB.Menu m_Copy_Hex 
            Caption         =   "Hex"
         End
         Begin VB.Menu m_Copy_Disassembly 
            Caption         =   "反汇编"
         End
      End
      Begin VB.Menu m_Bin 
         Caption         =   "编辑"
      End
      Begin VB.Menu m_2 
         Caption         =   "-"
      End
      Begin VB.Menu m_Assemble 
         Caption         =   "汇编"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Boolean

Dim Dev As New 驱动
Sub AddListItems(str0$, str1$, str2$)
Dim ListItems1 As ListSubItems
Dim ListItems2 As ListSubItem
str2$ = LCase(str2$)
Set ListItems1 = ListView1.ListItems.Add(, , str0$)
          ListItems1.Add , , str1$
Select Case Left$(str2$, 3)
     Case "cal"
     Set ListItems2 = ListItems1.Add(, , str2$)
          ListItems2.ForeColor = &HA50000: ListItems2.Bold = True: Exit Sub
     Case "ret"
     Set ListItems2 = ListItems1.Add(, , str2$)
          ListItems2.ForeColor = &HA50000: Exit Sub
End Select
If Left$(str2$, 1) = "j" Then
     If Left$(str2$, 4) <> "jmp " Then
     Set ListItems2 = ListItems1.Add(, , str2$)
          ListItems2.ForeColor = RGB(&HCE, &H82, 0): Exit Sub
     Else
     Set ListItems2 = ListItems1.Add(, , str2$)
          ListItems2.ForeColor = &H84: Exit Sub
     End If
End If
     ListItems1.Add , , str2$
End Sub

Private Sub Form_Load()
Dim Proc As New 进程
Proc.提升权限
Dim SysFile() As Byte
     SysFile = LoadResData(110, "HDADA")
     Open App.Path & "\hf.sys" For Binary As #1
     Put #1, , SysFile
     Close #1
If Dir(App.Path & "\hf.sys") = "" Then MsgBox "打开失败！请联系作者。", 32, "QQ:958570606": End

Dev.删除服务 "HanfY"
If Dir(App.Path & "\hf.sys") = "" Then MsgBox "出错退出！", vbExclamation: End
If (Dev.加载驱动(App.Path & "\hf.sys", "HanfY") <= 0) Then
MsgBox "加载驱动失败！", 48, "提示"
End
End If
Kill App.Path & "\hf.sys"
End Sub

Private Sub Form_Resize()
If IsIconic(Form1.hwnd) Then Exit Sub
If Height <= 700 + Text1.Height + 285 Then Height = 700 + Text1.Height + 285
ListView1.Height = IIf(Height > 700, Height - 700 - Text1.Height - 200, 700)
ListView1.Width = IIf(Width > 300, Width - 300, 300)
Text1.Top = ListView1.Top + ListView1.Height + 0
Text1.Width = IIf(Width > 300, Width - 300, 300)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dev.卸载驱动 "HanfY", True
End Sub

Private Sub ListView1_DblClick()
m_Assemble_Click
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu m_m
End If
End Sub

Private Sub m_Assemble_Click()
If ListView1.ListItems.Count <= 0 Then Exit Sub
If ListView1.SelectedItem.Selected Then
Dialog1.Show vbModal
End If
End Sub

Private Sub m_Bin_Click()
If ListView1.ListItems.Count <= 0 Then Exit Sub
If ListView1.SelectedItem.Selected Then
Dialog2.Show vbModal
End If
End Sub

Private Sub m_Copy_Addr_Click()
Dim strClip As String
Clipboard.Clear
strClip = ListView1.SelectedItem.Text
Clipboard.SetText strClip
End Sub

Private Sub m_Copy_Disassembly_Click()
If ListView1.ListItems.Count <= 0 Then Exit Sub
Dim clipStr As String
Dim n As Long
clipStr = ""
Clipboard.Clear
For n = 1 To ListView1.ListItems.Count
     If ListView1.ListItems(n).Selected Then
     clipStr = clipStr & ListView1.ListItems(n).ListSubItems(2) & vbCrLf
     End If
Next
Clipboard.SetText clipStr
End Sub

Private Sub m_Copy_Hex_Click()
If ListView1.ListItems.Count <= 0 Then Exit Sub
Dim clipStr As String
Dim n As Long
clipStr = ""
Clipboard.Clear
For n = 1 To ListView1.ListItems.Count
     If ListView1.ListItems(n).Selected Then
     clipStr = clipStr & ListView1.ListItems(n).ListSubItems(1) & vbCrLf
     End If
Next
Clipboard.SetText clipStr
End Sub

Private Sub m_Disassembly_Click()
Dialog.Show vbModal
End Sub


Private Sub m_Goto_Click()
Dim tempStrAry() As String
Dim tempVar As Long
If ListView1.ListItems.Count <= 0 Then Exit Sub
tempStrAry = Split(ListView1.SelectedItem.SubItems(2), " ", , vbTextCompare)
If SafeArrayGetDim(tempStrAry()) = False Then Exit Sub
If UBound(tempStrAry) >= 1 Then
     StrToIntExA "0x" & tempStrAry(UBound(tempStrAry)), 1, tempVar
     If tempVar = 0 Or (tempVar > &H0 And tempVar <= MEM_4MB_PAGES) Then
     Exit Sub
     Else
     If lOldSize <= 0 Then lOldSize = &H200
     StartDisassembly tempVar, lOldSize
     End If
End If

End Sub '作者：Leaf
