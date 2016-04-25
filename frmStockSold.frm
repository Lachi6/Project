VERSION 5.00
Begin VB.Form frmStockSold 
   Caption         =   "Stock Sold"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   Picture         =   "frmStockSold.frx":0000
   ScaleHeight     =   4620
   ScaleWidth      =   8775
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H8000000E&
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   4800
      Picture         =   "frmStockSold.frx":9CAA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   975
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
      Left            =   6240
      Picture         =   "frmStockSold.frx":A8EC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtQtySol 
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
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   120
      Top             =   600
      Width           =   7935
   End
   Begin VB.Label lblQunatityRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Sold"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   2175
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Width           =   4935
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
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblProductName 
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
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblProductId 
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
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmStockSold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
frmStockOut.grdStockOut.TextMatrix(frmStockOut.grdStockOut.Row, 4) = txtQtySol.Text
Unload Me
frmStockOut.cmdNew.Enabled = True

End Sub

Private Sub txtQtyRec_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
End Sub

