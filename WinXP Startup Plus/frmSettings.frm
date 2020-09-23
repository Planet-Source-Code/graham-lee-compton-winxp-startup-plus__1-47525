VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[: Settings :]"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5130
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chk 
      Caption         =   "Systray on minimise"
      Height          =   285
      Index           =   1
      Left            =   2070
      TabIndex        =   8
      Top             =   720
      Width           =   2625
   End
   Begin VB.CheckBox chk 
      Caption         =   "Fade In/Out (Win2k/Xp Only)"
      Height          =   285
      Index           =   4
      Left            =   90
      TabIndex        =   7
      Top             =   1260
      Width           =   2895
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Save"
      Height          =   465
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   1620
      Width           =   1185
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Close"
      Height          =   465
      Index           =   0
      Left            =   3780
      TabIndex        =   5
      Top             =   1620
      Width           =   1185
   End
   Begin VB.TextBox txtLogo 
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Text            =   "$me\XpS+.sys\logo.jpg"
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CheckBox chk 
      Caption         =   "Run onConn only Once!"
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   990
      Width           =   2355
   End
   Begin VB.CheckBox chk 
      Caption         =   "Minimise on Start"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   720
      Width           =   2625
   End
   Begin VB.CheckBox chk 
      Caption         =   "Save Form Settings on Exit"
      Height          =   285
      Index           =   3
      Left            =   2430
      TabIndex        =   0
      Top             =   990
      Width           =   2625
   End
   Begin VB.Image imgBrowse 
      Height          =   480
      Left            =   4500
      Picture         =   "frmSettings.frx":0000
      Top             =   270
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Logo path"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
