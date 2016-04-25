VERSION 5.00
Begin VB.Form frmShortcut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shortcut Keys"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmShortcut.frx":0000
   ScaleHeight     =   6045
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   600
      Picture         =   "frmShortcut.frx":9CAA
      ScaleHeight     =   4755
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   360
      Width           =   3255
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000E&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "F2"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Customer"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard Shortcuts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Log Off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock Application"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "About Us"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + O"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + C"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + L"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "F5"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "F6"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + S"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + D"
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + U"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   0
         Picture         =   "frmShortcut.frx":13954
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
