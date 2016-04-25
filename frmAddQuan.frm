VERSION 5.00
Begin VB.Form frmAddQuanP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Quantity"
   ClientHeight    =   3390
   ClientLeft      =   4035
   ClientTop       =   3600
   ClientWidth     =   6945
   Icon            =   "frmAddQuan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmAddQuan.frx":23D2
   ScaleHeight     =   3390
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtQuan 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
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
      Left            =   3840
      Picture         =   "frmAddQuan.frx":C07C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H8000000E&
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
      Left            =   2160
      Picture         =   "frmAddQuan.frx":CCBE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   720
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   720
      Top             =   120
      Width           =   5775
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
      Left            =   1080
      TabIndex        =   7
      Top             =   480
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
      Left            =   3000
      TabIndex        =   6
      Top             =   360
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
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   1575
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
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   3015
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "frmAddQuanP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rspo As New ADODB.Recordset


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Set rspo = New ADODB.Recordset
rspo.Open "select * from Model where ModelNo='" & scAd & "'", con, adOpenKeyset, adLockPessimistic
If rspo.RecordCount > 0 Then
    If IsNumeric(txtQuan.Text) = True Then
    qtyAd = txtQuan.Text
    With frmPurchaseOrd.lvwOrder
      .ListItems.Add , , scAd
      .ListItems(.ListItems.Count).ListSubItems.Add , , itmAd
      .ListItems(.ListItems.Count).ListSubItems.Add , , qtyAd
      .ListItems(.ListItems.Count).ListSubItems.Add , , rspo.Fields(2).Value
      .ListItems(.ListItems.Count).ListSubItems.Add , , sup
    End With
    Else
    txtQuan.Text = ""
    MsgBox "Please enter proper quntity.", vbOKOnly + vbInformation, "Error"
    Exit Sub
    End If
    
 End If
 Set rspo = Nothing
 scAd = ""
 itmAd = ""
 qtyAd = ""
 Unload Me
 frmPurchaseOrd.cmdAdd.Enabled = True
 
End Sub

Private Sub txtQuan_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
End Sub
