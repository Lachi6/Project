VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add User"
   ClientHeight    =   7635
   ClientLeft      =   4095
   ClientTop       =   3660
   ClientWidth     =   7650
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAddUser.frx":0E42
   ScaleHeight     =   7635
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPh 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.ComboBox cboStatus 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmAddUser.frx":AAEC
      Left            =   3120
      List            =   "frmAddUser.frx":AAF6
      TabIndex        =   6
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtPass1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox txtPass2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3720
      Width           =   3135
   End
   Begin VB.ComboBox cboAccess 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmAddUser.frx":AB0B
      Left            =   3120
      List            =   "frmAddUser.frx":AB15
      TabIndex        =   5
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000E&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      Picture         =   "frmAddUser.frx":AB2E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000E&
      Caption         =   "&Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Picture         =   "frmAddUser.frx":B770
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   4695
      Left            =   720
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   720
      Top             =   5520
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add New User"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   150
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2520
      Picture         =   "frmAddUser.frx":C3B2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblAccess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Access Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   960
      TabIndex        =   13
      Top             =   4320
      Width           =   1350
   End
   Begin VB.Label lblPass2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retype Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   960
      TabIndex        =   12
      Top             =   3840
      Width           =   1875
   End
   Begin VB.Label lblPass1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   960
      TabIndex        =   11
      Top             =   3360
      Width           =   1050
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   960
      TabIndex        =   10
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   960
      TabIndex        =   9
      Top             =   1440
      Width           =   795
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
blank
Unload Me

End Sub

Private Sub cmdSave_Click()
  If txtUserID.Text = "" Or txtName.Text = "" Or txtUserName.Text = "" Or txtPass1.Text = "" Or txtPh.Text = "" Or txtPass2.Text = "" Or cboAccess.ListIndex = -1 Then
   MsgBox "Please enter required information", vbOKOnly, "Validation"

   Exit Sub
  End If

  If txtPass1.Text <> txtPass2.Text Then
   MsgBox "Both Passwords does not match", vbOKOnly, "Validation"
  Exit Sub

ElseIf txtPass1.Text = txtPass2.Text Then

With rs_Password
.AddNew
.Fields(0) = txtUserID.Text
.Fields(1) = txtName.Text
.Fields(2) = txtUserName.Text
.Fields(3) = txtPass1.Text
.Fields(4) = cboAccess.Text
.Fields(5) = cboStatus.Text
.Fields(6) = txtPh.Text
.Update

MsgBox "User Successfully Added", vbOKOnly, "Success"
End With
End If
blank
Call createID
txtName.SetFocus

frmLogin.Show
End Sub

Private Sub Form_Load()
frmAddUser.Top = 350
frmAddUser.Left = 4100

Call Connect
Call createID

End Sub
Sub blank()
txtUserID.Text = ""
txtName.Text = ""
txtUserName.Text = ""
txtPass1.Text = ""
txtPass2.Text = ""
cboAccess.ListIndex = -1
cboStatus.ListIndex = -1

End Sub
Sub createID()
Dim lastno As Long, newno As Long
'check if there are record in the file
With rs_Password
 If .BOF = True And .EOF = True Then
 lastno = 0
 Else
 .MoveLast
 lastno = CLng(Mid(.Fields(0), 2, 2))
 End If
 'generate new no
  newno = lastno + 1
  txtUserID.Text = "U" & newno
End With
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Call CheckspChar(KeyAscii)
End Sub


Private Sub txtPh_KeyPress(KeyAscii As Integer)
Call DonotAllowSpChar(KeyAscii)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
Call CheckspChar(KeyAscii)
End Sub
