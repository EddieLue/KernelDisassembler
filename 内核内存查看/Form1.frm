VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "内核内存监视器"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9840
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   4935
      TabIndex        =   8
      Top             =   240
      Width           =   750
   End
   Begin VB.ComboBox Combo_Type 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   3120
      List            =   "Form1.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8505
      Top             =   5355
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   4710
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "增加"
      Height          =   375
      Left            =   8940
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text_MemData 
      Height          =   300
      Left            =   6510
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text_Address 
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3885
      Left            =   -15
      TabIndex        =   0
      Top             =   720
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6853
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "地址"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "类型"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "内容"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "长度:"
      Height          =   270
      Left            =   4410
      TabIndex        =   9
      Top             =   285
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "内容:"
      Height          =   255
      Left            =   6030
      TabIndex        =   5
      Top             =   285
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "地址:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   285
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim NodeItem As ListItem
    If (Len(Text_Address) <= 0 Or Len(Text_MemData) <= 0) Then Exit Sub
    Set NodeItem = ListView1.ListItems.Add(, , Text_Address)
    NodeItem.ListSubItems.Add , , Combo_Type.Text
    NodeItem.ListSubItems.Add , , Text_MemData
End Sub

Private Sub Form_Load()
Dim proc As New 进程
    Combo_Type.ListIndex = 2
    proc.提升权限
    If Dir(App.Path & "\hf.sys") = "" Then MsgBox "出错退出！", vbExclamation: End
    
    If (Dev.加载驱动(App.Path & "\hf.sys", "HanfMem") <= 0) Then
        MsgBox "加载驱动失败！", 48, "提示"
        End
    End If
    
    'Kill App.Path & "\hf.sys"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dev.卸载驱动 True
End Sub

Private Sub Form_Resize()
    ListView1.Width = Width - 5
    ListView1.Height = Height - StatusBar1.Height - 200
    ListView1.ColumnHeaders(3).Width = ListView1.Width - (ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(1).Width)
End Sub



Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = vbKeyRButton Then
      ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If
End Sub

Private Sub Text_Address_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
Dim nAddr As Long
Dim nSize As Long
Dim nAddrSize(0 To 1) As Long
Dim bOutBuf() As Byte

    If (KeyCode = vbKeyReturn) Then
        If (StrToIntExA(Text_Address, 1, nAddr) = False) Then Exit Sub
        If (nAddr = 0) Then Exit Sub
        
        nSize = Abs(Combo_Type.ItemData(Combo_Type.ListIndex))
        ReDim bOutBuf(0 To nSize + 4)
        
        nAddrSize(0) = nAddr
        nAddrSize(1) = nSize
        
        ReadMem nAddrSize(), bOutBuf(), nSize + 4
        
         MsgBox DbgPrintArray(bOutBuf(), nSize + 4)
    End If
    
End Sub
