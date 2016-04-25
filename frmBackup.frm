VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Database"
   ClientHeight    =   3765
   ClientLeft      =   4095
   ClientTop       =   3660
   ClientWidth     =   8685
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmBackup.frx":23D2
   ScaleHeight     =   3765
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000E&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Picture         =   "frmBackup.frx":C07C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdpath 
      BackColor       =   &H8000000E&
      Caption         =   "&Select Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Picture         =   "frmBackup.frx":CCBE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdbackup 
      BackColor       =   &H8000000E&
      Caption         =   "&Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Picture         =   "frmBackup.frx":D900
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtpath 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Path for database backup"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   240
      Top             =   840
      Width           =   8175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Backup Of Database"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   170
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   2040
      Picture         =   "frmBackup.frx":E542
      Stretch         =   -1  'True
      Top             =   165
      Width           =   4695
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdbackup_Click()
Dim spath As String
Dim dpath As String

spath = App.Path + "\Mobile.mdb"
If txtpath.Text <> "" Then
dpath = txtpath.Text
FileCopy spath, dpath
MsgBox "Database backup done successfully.", vbOKOnly, "Success"

If cmdpath.Visible = False Then
cmdpath.Visible = True
End If
If cmdbackup.Visible = True Then
cmdbackup.Visible = False
End If

Else
MsgBox "Destination path is not selected", vbOKOnly + vbInformation, "Error"
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPath_Click()
With CommonDialog1
.Filter = "Microsoft Database Files|*.mdb"
.ShowSave
End With
txtpath.Text = CommonDialog1.FileName

If txtpath.Text <> "" Then
MsgBox "Backup file is created.", vbOKOnly, "Success"
 If cmdbackup.Visible = False Then
 cmdbackup.Visible = True
 End If
 
 If cmdpath.Visible = True Then
 cmdpath.Visible = False
 End If
End If
End Sub

Private Sub Form_Load()
con.Close

frmBackup.Top = 350
frmBackup.Left = 4100
txtpath.Enabled = False
cmdbackup.Visible = False
End Sub



