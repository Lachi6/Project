VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStockOut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Out"
   ClientHeight    =   6285
   ClientLeft      =   4230
   ClientTop       =   3660
   ClientWidth     =   8940
   Icon            =   "frmStockOut.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmStockOut.frx":23D2
   ScaleHeight     =   6285
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDateRec 
      BackColor       =   &H80000018&
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
      Left            =   6480
      TabIndex        =   13
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtCustName 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtDateOrd 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cboSalesOrd 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H8000000E&
      Caption         =   "New"
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
      Left            =   4920
      Picture         =   "frmStockOut.frx":C07C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Remove sold product from stock"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000E&
      Cancel          =   -1  'True
      Caption         =   "&Close"
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
      Left            =   6480
      Picture         =   "frmStockOut.frx":CCBE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close this window"
      Top             =   5280
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdStockOut 
      Height          =   2055
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorSel    =   12640511
      ForeColorSel    =   16744703
      BackColorBkg    =   -2147483624
      GridColor       =   4210752
      GridColorFixed  =   49152
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Out"
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
      Left            =   3360
      TabIndex        =   12
      Top             =   195
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3120
      Picture         =   "frmStockOut.frx":D900
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblStockOut 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblCustName 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblDateRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Recieved"
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
      Left            =   4800
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblDateOrd 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Ordered"
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
      Left            =   4800
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblSONo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order No."
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
   Begin VB.Label lblStockOutNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Out No."
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
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   240
      Top             =   960
      Width           =   8415
   End
End
Attribute VB_Name = "frmStockOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsStockout As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub cboSalesOrd_Click()
rsStockout.Open "select * from S_Order where SO_No='" & cboSalesOrd.Text & "'", con, adOpenKeyset, adLockOptimistic
grdStockOut.Rows = 1
Do While Not rsStockout.EOF
   txtCustName.Text = rsStockout.Fields!Customer_Name
   txtDateOrd.Text = rsStockout.Fields![SO_Date_Order]
   txtDateRec.Text = Date
   grdStockOut.AddItem rsStockout.Fields![Product_ID]
   grdStockOut.TextMatrix(grdStockOut.Rows - 1, 3) = rsStockout.Fields![Quantity]
   rs2.Open "select * from Products where Product_ID='" & rsStockout.Fields![Product_ID] & "'", con, adOpenKeyset, adLockPessimistic
   If rs2.RecordCount > 0 Then
   grdStockOut.TextMatrix(grdStockOut.Rows - 1, 1) = rs2.Fields![Product_Name]
   grdStockOut.TextMatrix(grdStockOut.Rows - 1, 2) = rs2.Fields![Unit_Price]
   
   End If
   rs2.Close
   rsStockout.MoveNext
 Loop
 Set rsStockout = Nothing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
If cmdNew.Caption = "New" Then
  cboSalesOrd.Enabled = True
  grdStockOut.Enabled = True
  
  cboSalesOrd.Text = ""
  grdStockOut.Row = 0
  'put all the customer name in the combobox
rsStockout.Open "select SO_No from S_Order", con, adOpenKeyset, adLockOptimistic
   Do While Not rsStockout.EOF
      cboSalesOrd.AddItem rsStockout.Fields!SO_No
      rsStockout.MoveNext
   Loop
 Set rsStockout = Nothing
 Call createID
 cmdNew.Caption = "Save"
 cmdNew.Enabled = False
 cboSalesOrd.SetFocus
Else
'put all stock in Stock In table
 rsStockout.Open "select * from Stock_out", con, adOpenKeyset, adLockOptimistic
  For i = 1 To grdStockOut.Rows - 1
     rsStockout.AddNew
     rsStockout.Fields(0) = lblStockOut.Caption
     rsStockout.Fields(1) = cboSalesOrd.Text
     rsStockout.Fields(2) = Date
     rsStockout.Fields(3) = grdStockOut.TextMatrix(i, 0)
     rsStockout.Fields(4) = grdStockOut.TextMatrix(i, 4)
     rsStockout.Update
     'add the unit in stock to the Products table
     rs2.Open "select * from Products where Product_ID='" & grdStockOut.TextMatrix(i, 0) & "'", con, adOpenKeyset, adLockOptimistic
     If rs2.RecordCount > 0 Then
       rs2.Fields![Unit_In_Stock] = Val(rs2.Fields![Unit_In_Stock]) - Val(grdStockOut.TextMatrix(i, 4))
       rs2.Update
    End If
       Set rs2 = Nothing
    Next i
    Set rsStockout = Nothing
    
    'delete SO that already stock out
    rsStockout.Open "delete from S_Order where SO_No='" & cboSalesOrd.Text & "'", con, adOpenKeyset, adLockOptimistic
    MsgBox "Stock successfully removed.", vbOKOnly + vbInformation, "Success"
    cmdNew.Caption = "New"
    
    'clear and disable the components
    cboSalesOrd.clear
    cboSalesOrd.Enabled = False
    grdStockOut.Enabled = False
    cboSalesOrd.Text = ""
    txtDateOrd.Text = ""
    txtDateRec.Text = ""
    txtCustName.Text = ""
    grdStockOut.Rows = 1
    lblStockOut.Caption = ""
    Set rsStockout = Nothing
 End If
 'cmdNew.Enabled = False
End Sub

Private Sub Form_Load()
Dim i As Integer
i = i + 1

grdStockOut.TextMatrix(0, 0) = "Product ID"
grdStockOut.TextMatrix(0, 1) = "Name"
grdStockOut.TextMatrix(0, 2) = "amount"
grdStockOut.TextMatrix(0, 3) = "quantity"
grdStockOut.TextMatrix(0, 4) = "sold"

With grdStockOut
   .ColWidth(0) = 1300
   .ColWidth(1) = 3300
   .ColWidth(2) = 1100
   .ColWidth(3) = 1100
   .ColWidth(4) = 1100
End With
Call Connect

frmStockOut.Top = 350
frmStockOut.Left = 4100
End Sub
Sub createID()
Dim lastno As Long, newno As Long
Set rsStockout = New ADODB.Recordset
rsStockout.Open "select * from Stock_out", con, adOpenDynamic, adLockOptimistic

With rsStockout
 If .BOF = True And .EOF = True Then
 lastno = 0
 Else
 .MoveLast
 lastno = CLng(Mid(.Fields(0), 3, 2))
 End If
 
  newno = lastno + 1
   lblStockOut.Caption = "SO" & newno
End With
Set rsStockout = Nothing

End Sub

Private Sub grdStockOut_Click()
frmStockSold.lblProdID.Caption = grdStockOut.TextMatrix(grdStockOut.Row, 0)
frmStockSold.lblProdName.Caption = grdStockOut.TextMatrix(grdStockOut.Row, 1)
frmStockSold.Show vbModal

End Sub


