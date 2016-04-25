VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Us"
   ClientHeight    =   4245
   ClientLeft      =   3735
   ClientTop       =   3165
   ClientWidth     =   7530
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H8000000E&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      Picture         =   "frmAbout.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1084
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   6
      Top             =   3720
      Width           =   6495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   0
      X2              =   7560
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Vinay Kumar Singh"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vikas Shukla"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   12285
      Left            =   0
      Picture         =   "frmAbout.frx":112E
      Top             =   1560
      Width           =   15360
   End
   Begin VB.Image Image2 
      Height          =   1230
      Left            =   240
      Picture         =   "frmAbout.frx":99DE
      Top             =   240
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   1440
      Picture         =   "frmAbout.frx":A60A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdok_Click()
Unload Me

End Sub

Private Sub Form_Load()
frmAbout.Top = 1500
frmAbout.Left = 3990
End Sub
