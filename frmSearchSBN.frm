VERSION 5.00
Begin VB.Form frmSearchSBN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Searched Supplier "
   ClientHeight    =   7305
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmSearchSBN.frx":0000
   ScaleHeight     =   7305
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H8000000B&
      Caption         =   "&OK"
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
      Left            =   3480
      Picture         =   "frmSearchSBN.frx":9CAA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox txtSupPhone 
      BackColor       =   &H80000018&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox txtSupAddr 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox txtSupMail 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox txtSupName 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtSupID 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtdate 
      BackColor       =   &H80000018&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers"
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
      Left            =   2760
      TabIndex        =   15
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This is the Supplier to be searched!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   5760
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers Details"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblSupMail 
      BackStyle       =   0  'Transparent
      Caption         =   "e-Mail"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblSupPhone 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone / Mo. No."
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
      Left            =   840
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label LblSupAddr 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblSupName 
      BackStyle       =   0  'Transparent
      Caption         =   " Name"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblSupID 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   4575
      Left            =   600
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label lblSupDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2640
      Picture         =   "frmSearchSBN.frx":A8EC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmSearchSBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssupplier As New ADODB.Recordset
Dim rsaddsup As New ADODB.Recordset
Dim name1 As String
Dim n As Integer
Private Sub cmdok_Click()
Unload Me
End Sub
Private Sub Form_Load()
frmSearchSBN.Top = 350
frmSearchSBN.Left = 4100

txtSupID.locked = True
txtSupName.locked = True
txtSupAddr.locked = True
txtSupPhone.locked = True
txtSupMail.locked = True
txtdate.locked = True
name1 = InputBox("Please enter the name of supplier to be searched.", vbOKOnly)
Call Connect
Set rssupplier = New ADODB.Recordset
rssupplier.Open "select * from Suppliers", con, adOpenDynamic, adLockOptimistic

With rssupplier
While Not .EOF
    If .Fields(1).Value = name1 Then
        Call showdata
        Exit Sub
        Else
        Label3.Caption = "Supplier Not Found In system"
     End If
.MoveNext
Wend

End With
End Sub

Sub showdata()

txtSupID.Text = rssupplier.Fields(0)
txtSupName.Text = rssupplier.Fields(1)
txtSupAddr.Text = rssupplier.Fields(2)
txtSupPhone.Text = rssupplier.Fields(3)
txtSupMail.Text = rssupplier.Fields(4)
txtdate.Text = rssupplier.Fields(5)
Label3.Caption = "This is the Supplier to be searched!!!"
End Sub
