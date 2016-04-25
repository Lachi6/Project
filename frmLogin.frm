VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Mobile Shoppee Management Login"
   ClientHeight    =   3675
   ClientLeft      =   4050
   ClientTop       =   3090
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFF00&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboUserName 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmLogin.frx":C472
      Left            =   2400
      List            =   "frmLogin.frx":C474
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Image Image4 
      Height          =   1140
      Left            =   120
      Picture         =   "frmLogin.frx":C476
      Top             =   2400
      Width           =   1710
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   1560
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   5280
      Picture         =   "frmLogin.frx":CD11
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter User Name and Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "frmLogin.frx":12F6B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   6855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                                                                                       
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdLogin_Click()
If cboUserName.ListIndex = -1 Then
 MsgBox "Please select a user name", vbOKOnly + vbInformation, "Error"
 Exit Sub
End If

UserName = cboUserName.List(cboUserName.ListIndex)
With rs_Password
   .MoveFirst
  While Not .EOF
    If UserName = .Fields(2).Value And txtPass.Text = .Fields(3).Value Then
      Rights = .Fields(4).Value
      Status = .Fields(5).Value
       If Status = "Blocked" Then
       MsgBox "Selected user is Blocked.", vbOKOnly + vbInformation, "Error"
       Exit Sub
       ElseIf Rights = "User" Then
       Unload Me
       frmSale.Show
       Else
           Unload Me
           MDIForm1.Show
           frmSideBar.Show
       End If
     
    
      ElseIf UserName = .Fields(2).Value And txtPass.Text <> .Fields(3).Value Then
      MsgBox "Incorrect Password", vbOKOnly + vbInformation, "Error"
      Exit Sub
       
    End If
     
    .MoveNext
    Wend

End With

If Rights = "Administrator" Then
MDIForm1.StatusBar1.Panels(3).Text = UserName
frmSideBar.lblUserName.Caption = UserName
frmSideBar.lblTime.Caption = Time

ElseIf Rights = "User" Then
MDIForm1.StatusBar1.Panels(3).Text = UserName
MDIForm1.mnuSearch.Enabled = False
MDIForm1.mnuReport.Enabled = False
MDIForm1.mnuFile.Enabled = False
MDIForm1.mnuAdmin.Enabled = False
MDIForm1.Toolbar1.Enabled = False
frmSale.Show
'frmSidebar.lblUserName.Caption = UserName
'frmSidebar.lblTime.Caption = Time

End If


End Sub

Private Sub Form_Load()

Call Connect

With rs_Password
.MoveFirst
While Not .EOF
cboUserName.AddItem .Fields(2).Value
.MoveNext
Wend
End With

End Sub



