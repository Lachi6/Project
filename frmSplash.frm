VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4575
   ClientLeft      =   3675
   ClientTop       =   2805
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.Timer Timer1 
         Interval        =   20
         Left            =   240
         Top             =   3120
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   3000
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   1800
         Left            =   0
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label lblPer 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Image Image3 
         Height          =   10005
         Left            =   -5400
         Picture         =   "frmSplash.frx":0E07
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   15360
      End
      Begin VB.Image Image2 
         Height          =   1695
         Left            =   4440
         Picture         =   "frmSplash.frx":96B7
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2715
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
lblPer.Caption = ProgressBar1.Value
If ProgressBar1.Value <= 20 Then
    lblStatus.Caption = "Initializing.."
    ElseIf ProgressBar1.Value <= 50 Then
    lblStatus.Caption = "Loading components..."
    ElseIf ProgressBar1.Value <= 85 Then
    lblStatus.Caption = "Integrating Database...."
    ElseIf ProgressBar1.Value <= 99 Then
    lblStatus.Caption = "Starting Application....."
    
 End If
    
    If ProgressBar1.Value = 100 Then
       Unload Me
        frmAddUser.Show
        
        'Load frmLogin
        'frmLogin.Show
        End If
        

End Sub

