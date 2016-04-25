VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Users"
   ClientHeight    =   4785
   ClientLeft      =   4095
   ClientTop       =   3660
   ClientWidth     =   8190
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmUser.frx":0E42
   ScaleHeight     =   4785
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
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
      ItemData        =   "frmUser.frx":AAEC
      Left            =   4200
      List            =   "frmUser.frx":AAF6
      TabIndex        =   14
      Top             =   2520
      Width           =   2655
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
      ItemData        =   "frmUser.frx":AB0B
      Left            =   4200
      List            =   "frmUser.frx":AB15
      TabIndex        =   13
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H8000000E&
      Caption         =   "&Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Picture         =   "frmUser.frx":AB2E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Save the modification"
      Top             =   3360
      Width           =   975
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
      Height          =   360
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
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
      Height          =   360
      Left            =   4200
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
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
      Left            =   6120
      Picture         =   "frmUser.frx":B770
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000E&
      Caption         =   "&Edit"
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
      Left            =   3720
      Picture         =   "frmUser.frx":C3B2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Edit a user"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H8000000E&
      Caption         =   "&Add"
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
      Left            =   2520
      Picture         =   "frmUser.frx":CFF4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Add a new user to record"
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape3 
      Height          =   3375
      Left            =   240
      Top             =   960
      Width           =   2055
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
      Left            =   2640
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   2400
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "All Users"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3000
      Picture         =   "frmUser.frx":DC36
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   2400
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label lblUserID 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblAccess 
      BackStyle       =   0  'Transparent
      Caption         =   "Access Rights"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
Me.Hide
frmAddUser.Show

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
If txtUserID.Text = "" Then
MsgBox "No record is selected", vbOKOnly + vbInformation, "Error"
Exit Sub
Else

Call unlocked
End If
cmdUpdate.Enabled = True

End Sub

Private Sub cmdUpdate_Click()
If rs_Password.BOF = True And rs_Password.EOF = True Then Exit Sub
On Error Resume Next
 If txtUserID.Text = "" Then
 MsgBox "Please select a record and click on Edit.", vbOKOnly + vbInformation, "Error"
  Exit Sub
 Else
 Dim rs_update As ADODB.Recordset
 Set rs_update = New ADODB.Recordset
 If rs_update.State = 1 Then rs_update.Close
 rs_update.Open "select * from Login where User_ID='" & txtUserID.Text & "'", con, adOpenKeyset, adLockOptimistic
  With rs_update
  .Fields(0) = txtUserID.Text
  .Fields(1) = txtName.Text
  .Fields(4) = cboAccess.Text
  .Fields(5) = cboStatus.Text
  End With
  rs_update.Update
  MsgBox "Record is updated successfully.", vbOKOnly + vbInformation, "Error"
  Call locked
 End If
  
  End Sub

Private Sub Form_Load()
frmUser.Top = 350
frmUser.Left = 4100
Call Connect
With rs_Password
.MoveFirst
While Not .EOF
List1.AddItem .Fields(2)
.MoveNext
Wend
End With


End Sub


Private Sub List1_Click()
locked

With rs_Password
.MoveFirst
While Not .EOF
If List1.List(List1.ListIndex) = .Fields(2) Then
txtUserID.Text = .Fields(0)
txtName.Text = .Fields(1)
cboAccess.Text = .Fields(4)
cboStatus.Text = .Fields(5)

End If
.MoveNext
Wend
End With

End Sub
Sub locked()
txtName.locked = True
cboAccess.locked = True
cboStatus.locked = True


End Sub
Sub unlocked()
txtName.locked = False
cboAccess.locked = False
cboStatus.locked = False

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Call CheckspChar(KeyAscii)
End Sub

