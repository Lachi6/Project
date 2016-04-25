VERSION 5.00
Begin VB.Form frmLocked 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application is locked"
   ClientHeight    =   3090
   ClientLeft      =   4095
   ClientTop       =   3660
   ClientWidth     =   6615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLocked.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmLocked.frx":23D2
   ScaleHeight     =   3090
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUnloack 
      BackColor       =   &H8000000E&
      Caption         =   "&Unlock"
      Default         =   -1  'True
      Height          =   855
      Left            =   2760
      Picture         =   "frmLocked.frx":C07C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter password to unlock application"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   360
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label lblLocked 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Locked"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2280
      Picture         =   "frmLocked.frx":CCBE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rslock As ADODB.Recordset

Private Sub cmdUnloack_Click()
If txtPass.Text = "" Then
MsgBox "Please enter password", vbOKOnly + vbInformation, "Error"
Exit Sub
ElseIf txtPass.Text = rslock("Password") Then
Unload Me
Else
MsgBox "Wrong password! Attemp to unlock is failed.", vbOKOnly + vbInformation, "Error"
Exit Sub
txtPass.SetFocus

End If

End Sub

Private Sub Form_Load()
frmLocked.Top = 2500
frmLocked.Left = 4200

Call Connect
Set rslock = New ADODB.Recordset
rslock.Open "Select * from Login where User_name='" & UserName & "'", con, adOpenKeyset, adLockOptimistic

End Sub
