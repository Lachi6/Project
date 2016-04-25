VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   3915
   ClientLeft      =   4035
   ClientTop       =   3600
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmChangePass.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
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
      ItemData        =   "frmChangePass.frx":9CAA
      Left            =   3240
      List            =   "frmChangePass.frx":9CAC
      TabIndex        =   9
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox T 
      BackColor       =   &H80000018&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox T 
      BackColor       =   &H80000018&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox T 
      BackColor       =   &H80000018&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H80000002&
      Caption         =   "Change"
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
      Left            =   1800
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000004&
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
      Left            =   4080
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*Username and Password is Case-sensitive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter User Name and Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   11
      Top             =   120
      Width           =   3975
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
      Left            =   1440
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
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
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OLD  PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NEW  PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   6360
      Picture         =   "frmChangePass.frx":9CAE
      Top             =   0
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   0
      Picture         =   "frmChangePass.frx":FF08
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsso As New ADODB.Recordset
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdChange_Click()
If cboUserName.ListIndex = -1 Then
 MsgBox "Please select a user name", vbOKOnly + vbInformation, "Error"
 Exit Sub
End If

UserName = cboUserName.List(cboUserName.ListIndex)
With rs_Password
  .MoveFirst
  While Not .EOF
    If UserName = .Fields(2).Value And T(1) = .Fields(3).Value Then
      Rights = .Fields(4).Value
      Status = .Fields(5).Value
    If Status = "Blocked" Then
       MsgBox "Selected user is Blocked.", vbOKOnly + vbInformation, "Error"
       Exit Sub
    ElseIf UserName = .Fields(2).Value And T(1) <> .Fields(3).Value Then
      MsgBox "Incorrect Password", vbOKOnly + vbInformation, "Error"
      Exit Sub
    ElseIf T(2).Text <> T(3).Text Then
        MsgBox "The NEW Password and CONFIRM Password mismatched!", vbOKOnly + vbExclamation, "Sorry!"
        Exit Sub
     
     Else
            .Fields(3).Value = T(3).Text
            .Update
            MsgBox "Password has been changed successfully", vbOKOnly, "Success"
            Unload Me
        End If
         
End If
.MoveNext
        Wend
  End With
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




