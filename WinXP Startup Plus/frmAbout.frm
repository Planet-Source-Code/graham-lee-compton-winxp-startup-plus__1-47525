VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " [:: About ::]"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6105
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lb 
      Alignment       =   2  'Center
      Caption         =   "GraZ"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   7
      Left            =   1260
      TabIndex        =   7
      Top             =   3420
      Width           =   1455
   End
   Begin VB.Label lb 
      Alignment       =   2  'Center
      Caption         =   "12/12/02"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   1260
      TabIndex        =   6
      Top             =   3150
      Width           =   1455
   End
   Begin VB.Label lb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Made By:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   3420
      Width           =   885
   End
   Begin VB.Label lb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Build Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   4
      Top             =   3150
      Width           =   1065
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   3795
      Left            =   2970
      Stretch         =   -1  'True
      Top             =   90
      Width           =   3075
   End
   Begin VB.Label lb 
      Caption         =   "Developed to provide a much better enviroment and more simplistic use than the original XpS."
      Height          =   825
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   900
      Width           =   2895
   End
   Begin VB.Label lb 
      Alignment       =   2  'Center
      Caption         =   "Codename: tw@ XpS+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   540
      Width           =   2805
   End
   Begin VB.Label lb 
      Alignment       =   2  'Center
      Caption         =   "WinXP Startup Manager Plus"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   2805
   End
   Begin VB.Label lb 
      Caption         =   $"frmAbout.frx":0000
      Height          =   1275
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
