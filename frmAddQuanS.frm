VERSION 5.00
Begin VB.Form frmAddQuanS 
   Caption         =   "Enter Quantity"
   ClientHeight    =   3360
   ClientLeft      =   4110
   ClientTop       =   3675
   ClientWidth     =   6870
   Icon            =   "frmAddQuanS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmAddQuanS.frx":23D2
   ScaleHeight     =   3360
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQuan 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H80000009&
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
      Left            =   2040
      Picture         =   "frmAddQuanS.frx":C07C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000009&
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
      Left            =   3840
      Picture         =   "frmAddQuanS.frx":CCBE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   720
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label lblQuantity 
      BackStyle       =   0  'Transparent
      Caption         =   "Provide Quantity"
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
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblProdName 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblPrdN 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
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
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblProdID 
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
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
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
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Left            =   720
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmAddQuanS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsso As New ADODB.Recordset
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Set rsso = New ADODB.Recordset
rsso.Open "select * from Products where Product_ID='" & scAd & "'", con, adOpenKeyset, adLockPessimistic
If rsso.RecordCount > 0 Then
    If IsNumeric(txtQuan.Text) = True Then
    qtyAd = txtQuan.Text
    With frmSalesOrd.lvwOrder
      .ListItems.Add , , scAd
      .ListItems(.ListItems.Count).ListSubItems.Add , , itmAd
      .ListItems(.ListItems.Count).ListSubItems.Add , , qtyAd
      .ListItems(.ListItems.Count).ListSubItems.Add , , rsso.Fields(2).Value
      .ListItems(.ListItems.Count).ListSubItems.Add , , rsso.Fields(3).Value
      .ListItems(.ListItems.Count).ListSubItems.Add , , rsso.Fields(4).Value
      .ListItems(.ListItems.Count).ListSubItems.Add , , rsso.Fields(5).Value
    End With
    Else
    txtQuan.Text = ""
    MsgBox "Please enter proper quntity.", vbOKOnly + vbInformation, "Error"
    Exit Sub
    End If
    
 End If
 Set rsso = Nothing
 scAd = ""
 itmAd = ""
 qtyAd = ""
 Unload Me
 frmSalesOrd.cmdAdd.Enabled = True
End Sub


Private Sub txtQuan_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
End Sub
