VERSION 5.00
Begin VB.Form frmStockRec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Recieved"
   ClientHeight    =   3225
   ClientLeft      =   4095
   ClientTop       =   3660
   ClientWidth     =   5940
   Icon            =   "frmStockRec.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmStockRec.frx":23D2
   ScaleHeight     =   3225
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtModelNo 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdAddIMEI 
      BackColor       =   &H80000018&
      Caption         =   "ADD IMEI"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmStockRec.frx":C07C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtQtyRec 
      BackColor       =   &H80000014&
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
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000E&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Picture         =   "frmStockRec.frx":C503
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H8000000E&
      Caption         =   "&OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Picture         =   "frmStockRec.frx":D145
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblProductId 
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
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblProductName 
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
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblProdName 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblQunatityRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Recieved"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Left            =   240
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmStockRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddIMEI_Click()
Dim i As Integer
Dim a() As Long
Call Connect
Set rsproduct = New ADODB.Recordset
rsproduct.Open "select * from IMEI", con, adOpenDynamic, adLockOptimistic
For i = 0 To txtQtyRec.Text - 1
    X = InputBox("Enter IMEI Number")
    rsproduct.AddNew
    rsproduct.Fields(0) = X
    rsproduct.Fields(1) = txtModelNo.Text
    rsproduct.Fields(2) = txtPrice.Text
    rsproduct.Update
Next
rsproduct.Close
cmdOk.Enabled = True
cmdOk.SetFocus
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()

frmStockIn.grdStockIn.TextMatrix(frmStockIn.grdStockIn.Row, 4) = txtQtyRec.Text
frmStockIn.grdStockIn.TextMatrix(frmStockIn.grdStockIn.Row, 2) = txtPrice.Text
Unload Me
frmStockIn.cmdNew.Enabled = True


End Sub


Private Sub Form_Load()
'txtQtyRec.SetFocus
End Sub



Private Sub txtPrice_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
End Sub

Private Sub txtQtyRec_Change()
cmdAddIMEI.Enabled = True
End Sub

Private Sub txtQtyRec_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
End Sub
