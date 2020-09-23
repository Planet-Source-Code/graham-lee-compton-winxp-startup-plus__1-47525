VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "WinXP Startup Manager PLUS+"
   ClientHeight    =   5145
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9150
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   5145
   ScaleWidth      =   9150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrOnline 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   450
      Top             =   0
   End
   Begin MSComDlg.CommonDialog cD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select a Program to add"
      Filter          =   "EXE Files|*.exe|All Files|*.*"
   End
   Begin VB.PictureBox pAdd 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9090
      TabIndex        =   1
      Top             =   4680
      Width           =   9150
      Begin VB.CommandButton cmd 
         Caption         =   "Add"
         Height          =   285
         Index           =   2
         Left            =   6750
         TabIndex        =   8
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   5490
         TabIndex        =   7
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Edit"
         Height          =   285
         Index           =   1
         Left            =   6120
         TabIndex        =   6
         Top             =   90
         Width           =   555
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   4140
         TabIndex        =   5
         Text            =   "Path"
         ToolTipText     =   "Path to file"
         Top             =   90
         Width           =   1275
      End
      Begin VB.ComboBox cO 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":0442
         Left            =   2970
         List            =   "frmMain.frx":044C
         TabIndex        =   4
         Text            =   "onStart"
         ToolTipText     =   "Execute on Start or Internet Connect"
         Top             =   90
         Width           =   1095
      End
      Begin VB.ComboBox cO 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":0461
         Left            =   1890
         List            =   "frmMain.frx":046B
         TabIndex        =   3
         Text            =   "Normal"
         ToolTipText     =   "How to view the program"
         Top             =   90
         Width           =   1005
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Text            =   "Name"
         ToolTipText     =   "Name the program"
         Top             =   90
         Width           =   1725
      End
   End
   Begin MSComctlLib.ListView lst 
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   1455
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Style"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Method"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSetup 
         Caption         =   "Settings"
         Begin VB.Menu mnuOnConn 
            Caption         =   "Run onConn once"
         End
         Begin VB.Menu mnuMinimise 
            Caption         =   "Minimise on start"
         End
         Begin VB.Menu mnuAdd 
            Caption         =   "Hide Add tab"
         End
         Begin VB.Menu mnudash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "Alpha blending"
         End
      End
      Begin VB.Menu mnuSaveSettings 
         Caption         =   "Save Settings"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuAuthor 
         Caption         =   "Author"
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Website"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Selected"
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   "mnu2"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sOnline As Boolean

Private Sub cmd_Click(Index As Integer)
If Index = 0 Then
 cD.ShowOpen
 If Not cD.FileTitle = "" Then
  txt(1).Text = cD.FileName
 End If
End If
If Index = 1 Then
 Dim sOk As Boolean
 If txt(0).Text = "" Then Beep: Exit Sub
 If Not cO(0).Text = "Normal" And Not cO(0).Text = "Hidden" Then Beep: Exit Sub
 If Not Dir(txt(1)) = "" Then
    With lst
     .ListItems.Item(.SelectedItem.Index).Text = txt(0).Text
     .ListItems.Item(.SelectedItem.Index).ListSubItems.Item(1).Text = cO(0).Text
     .ListItems.Item(.SelectedItem.Index).ListSubItems.Item(2).Text = cO(1).Text
     .ListItems.Item(.SelectedItem.Index).ListSubItems.Item(3).Text = txt(1).Text
    End With
 End If
End If
If Index = 2 Then
 If txt(0).Text = "" Then Beep: Exit Sub
 If Not cO(0).Text = "Normal" And Not cO(0).Text = "Hidden" Then Beep: Exit Sub
 If Not Dir(txt(1)) = "" Then
  For Z = 1 To lst.ListItems.Count
   If lst.ListItems.Item(Z).Text = txt(0).Text Then sOk = True
  Next Z
  If sOk = True Then
    MsgBox "Error: Duplication Name Found", vbCritical, "Error.."
   Else
    With lst
     .ListItems.Add , , txt(0).Text
     .ListItems.Item(.ListItems.Count).ListSubItems.Add , , cO(0).Text
     .ListItems.Item(.ListItems.Count).ListSubItems.Add , , cO(1).Text
     .ListItems.Item(.ListItems.Count).ListSubItems.Add , , txt(1).Text
     .ListItems.Item(.ListItems.Count).Checked = True
    End With
  End If
 End If
End If
End Sub

Private Sub Form_Load()
 sOnline = False
 tmrOnline.Enabled = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
If WindowState = vbMinimized Then
   Hide
   Refresh
   With nid
   .cbSize = Len(nid)
   .hwnd = hwnd
   .uId = vbNull
   .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   .uCallBackMessage = WM_MOUSEMOVE
   .hIcon = Me.Icon
   .szTip = Me.Caption & vbNullChar
   End With
   Shell_NotifyIcon NIM_ADD, nid
End If
If Not WindowState = vbMinimized Then
 If pAdd.Visible = True Then
   lst.Height = ScaleHeight - pAdd.Height
   cmd(2).Left = ScaleWidth - cmd(2).Width - 100
   cmd(1).Left = cmd(2).Left - cmd(1).Width - 100
   cmd(0).Left = cmd(1).Left - cmd(0).Width - 100
   txt(1).Width = cmd(0).Left - txt(1).Left - 100
 End If
 If pAdd.Visible = False Then
   lst.Height = ScaleHeight
 End If
End If
 lst.Width = ScaleWidth
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = vbMinimized Then
  Dim Sys As Long
  Sys = X / Screen.TwipsPerPixelX
  Select Case Sys
  Case WM_LBUTTONDOWN:
   Shell_NotifyIcon NIM_DELETE, nid
   Me.WindowState = vbNormal
   Me.Visible = True
   Me.Show
  Case WM_RBUTTONDOWN:
   PopupMenu mnu2
  End Select
End If
End Sub

Private Sub mnuAlpha_Click()
 If mnuAlpha.Checked = True Then mnuAlpha.Checked = False Else mnuAlpha.Checked = True
End Sub

Private Sub mnuOnConn_Click()
 If mnuOnConn.Checked = True Then mnuOnConn.Checked = False Else mnuOnConn.Checked = True
End Sub

Private Sub mnuQuit_Click()
If mnuAlpha.Checked = True Then
  For Z = 255 To 0 Step -5
   Call MakeTransparent(frmMain.hwnd, Z)
  Next Z
  End
 Else
  End
End If
End Sub

Private Sub mnuRestore_Click()
 Shell_NotifyIcon NIM_DELETE, nid
 Me.WindowState = vbNormal
 Me.Visible = True
 Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Shell_NotifyIcon NIM_DELETE, nid
 End
End Sub

Private Sub lst_Click()
  If Not lst.SelectedItem Is Nothing Then
   txt(0).Text = lst.ListItems.Item(lst.SelectedItem.Index).Text
   cO(0).Text = lst.ListItems.Item(lst.SelectedItem.Index).ListSubItems.Item(1).Text
   cO(1).Text = lst.ListItems.Item(lst.SelectedItem.Index).ListSubItems.Item(2).Text
   txt(1).Text = lst.ListItems.Item(lst.SelectedItem.Index).ListSubItems.Item(3).Text
  End If
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then PopupMenu mnu
End Sub

Private Sub mnuAdd_Click()
On Error Resume Next
 If mnuAdd.Checked = True Then
   mnuAdd.Checked = False
   pAdd.Visible = True
   lst.Height = ScaleHeight - pAdd.Height
   cmd(2).Left = ScaleWidth - cmd(2).Width - 100
   cmd(1).Left = cmd(2).Left - cmd(1).Width - 100
   cmd(0).Left = cmd(1).Left - cmd(0).Width - 100
   txt(1).Width = cmd(0).Left - txt(1).Left - 100
  Else
   lst.Height = ScaleHeight
   mnuAdd.Checked = True
   pAdd.Visible = False
 End If
End Sub

Private Sub mnuAuthor_Click()
 frmAbout.Show
End Sub

Private Sub mnuMinimise_Click()
 If mnuMinimise.Checked = True Then mnuMinimise.Checked = False Else mnuMinimise.Checked = True
End Sub

Private Sub mnuRemove_Click()
 If Not lst.SelectedItem Is Nothing Then
  lst.ListItems.Remove (lst.SelectedItem.Index)
 End If
End Sub

Private Sub mnuSaveSettings_Click()
 WriteSettings
 Beep
End Sub

Private Sub mnuWebsite_Click()
 ShellExecute hwnd, "open", "www.grazc.com", vbNull, vbNull, vbNull
End Sub

Private Sub tmrOnline_Timer()
 If Online = True Then
  If sOnline = False Then
   For Z = 1 To lst.ListItems.Count
    If lst.ListItems.Item(Z).ListSubItems.Item(2).Text = "onConn" Then
      If lst.ListItems.Item(Z).Checked = True Then
       If Not Dir(lst.ListItems.Item(Z).ListSubItems.Item(3)) = "" Then
        If lst.ListItems.Item(Z).ListSubItems.Item(1).Text = "Normal" Then
         Shell lst.ListItems.Item(Z).ListSubItems.Item(3), vbNormalFocus
        End If
        If lst.ListItems.Item(Z).ListSubItems.Item(1).Text = "Hidden" Then
         Shell lst.ListItems.Item(Z).ListSubItems.Item(3), vbHide
        End If
       End If
      End If
     End If
     sOnline = True
   Next Z
  End If
  Else
  If mnuOnConn.Checked = True Then sOnline = False
 End If
End Sub
