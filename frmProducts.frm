VERSION 5.00
Begin VB.Form frmProducts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Products"
   ClientHeight    =   8985
   ClientLeft      =   4095
   ClientTop       =   3660
   ClientWidth     =   11805
   Icon            =   "frmProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmProducts.frx":29C12
   ScaleHeight     =   8985
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtprice 
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox txtBrand 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox txtProdStock 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   10
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtModelNo 
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox txtModelName 
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
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H8000000E&
      Caption         =   "&Last"
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
      Picture         =   "frmProducts.frx":338BC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Go to last record"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H8000000E&
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Picture         =   "frmProducts.frx":344FE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Go to previous record"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H8000000E&
      Caption         =   "&Next"
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
      Picture         =   "frmProducts.frx":35140
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Go to next record"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H8000000E&
      Caption         =   "&First"
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
      Left            =   1200
      Picture         =   "frmProducts.frx":35D82
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Go to first record"
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
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
      Left            =   1440
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Model No"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblProdStock 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit In Stock"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Products"
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
      Left            =   3960
      TabIndex        =   8
      Top             =   165
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   4575
      Left            =   840
      Top             =   840
      Width           =   6615
   End
   Begin VB.Label lblProdName 
      BackStyle       =   0  'Transparent
      Caption         =   "Model Name"
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
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3720
      Picture         =   "frmProducts.frx":369C4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Products Details"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      Height          =   1335
      Left            =   840
      Top             =   5640
      Width           =   6615
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsproduct As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private Sub Form_Load()

Call Connect
Set rsproduct = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
rsproduct.Open "select * from Model", con, adOpenDynamic, adLockOptimistic
rs1.Open "select * from IMEI where ModelNO='" & txtModelNo.Text & "'", con, adOpenKeyset, adLockPessimistic
stock = "Select count(IMEI_NO) From IMEI Where ModelNo like '" & txtModelNo.Text & "'"
showdata
End Sub
Sub showdata()

txtModelNo.Text = rsproduct.Fields(0)
txtModelName.Text = rsproduct.Fields(1)
txtBrand.Text = rsproduct.Fields(2)
txtprice.Text = rs1.Fields(2)
txtProdStock.Text = stock


                    
End Sub
Private Sub cmdFirst_Click()


On Error Resume Next
rsproduct.MoveFirst
  If rscustomer.BOF Then
      MsgBox "You are at first record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdLast_Click()


On Error Resume Next
rsproduct.MoveLast
  If rsproduct.EOF Then
      MsgBox "You are at last record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdNext_Click()

If rsproduct.BOF = True And rsproduct.EOF = True Then Exit Sub
On Error Resume Next
rsproduct.MoveNext
  If rsproduct.EOF Then
      MsgBox "You are at last record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdPrev_Click()


If rsproduct.BOF = True And rsproduct.EOF = True Then Exit Sub
On Error Resume Next
rsproduct.MovePrevious
  If rsproduct.BOF Then
      MsgBox "You are at first record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub
