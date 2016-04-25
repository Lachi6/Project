VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStockIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock In"
   ClientHeight    =   6390
   ClientLeft      =   4095
   ClientTop       =   3660
   ClientWidth     =   9255
   Icon            =   "frmStockIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmStockIn.frx":23D2
   ScaleHeight     =   6390
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSup 
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtDaterec 
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
      Left            =   6720
      TabIndex        =   13
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox cboPord 
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
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
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
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtSupName 
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
      Height          =   360
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
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
      Left            =   6360
      Picture         =   "frmStockIn.frx":C07C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close this window"
      Top             =   5400
      Width           =   975
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
      Picture         =   "frmStockIn.frx":CCBE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Add Product to stock"
      Top             =   5400
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdStockIn 
      Height          =   2055
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   7935
      _ExtentX        =   13996
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
      Caption         =   "Stock In"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   200
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3120
      Picture         =   "frmStockIn.frx":D900
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblStockInNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock In No."
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
      Left            =   600
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblPONo 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order No."
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
      Left            =   600
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
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
      Left            =   5040
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
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
      Left            =   5040
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblSupName 
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
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblStockIn 
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
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   4335
      Left            =   480
      Top             =   840
      Width           =   8295
   End
End
Attribute VB_Name = "frmStockIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsStockin As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Private Sub cboPord_Click()
rsStockin.Open "select * from P_Order where PO_No='" & cboPord.Text & "'", con, adOpenKeyset, adLockOptimistic
grdStockIn.Rows = 1
Do While Not rsStockin.EOF
   txtSupName.Text = rsStockin.Fields![BrandName]
   txtDateOrd.Text = rsStockin.Fields![PO_Date_Order]
   txtDaterec.Text = Date
   txtSup.Text = rsStockin.Fields![Supplier_name]
   grdStockIn.AddItem rsStockin.Fields![ModelNo]
   grdStockIn.TextMatrix(grdStockIn.Rows - 1, 3) = rsStockin.Fields![Quantity]
   rs2.Open "select * from Model where ModelNo='" & rsStockin.Fields![ModelNo] & "'", con, adOpenKeyset, adLockPessimistic
   If rs2.RecordCount > 0 Then
   grdStockIn.TextMatrix(grdStockIn.Rows - 1, 1) = rs2.Fields![ModelName]
'   grdStockIn.TextMatrix(grdStockIn.Rows - 1, 2) = rs2.Fields![Unit_Price]
   
   End If
   rs2.Close
   rsStockin.MoveNext
 Loop
 Set rsStockin = Nothing
 
   
  
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
If cmdNew.Caption = "New" Then
  cboPord.Enabled = True
  grdStockIn.Enabled = True
  
  cboPord.Text = ""
  grdStockIn.Row = 0
  'put all the supplier name in the combobox
rsStockin.Open "select PO_No from P_Order", con, adOpenKeyset, adLockOptimistic
   Do While Not rsStockin.EOF
      cboPord.AddItem rsStockin.Fields!PO_No
      rsStockin.MoveNext
   Loop
 Set rsStockin = Nothing
 Call createID
 cmdNew.Caption = "Save"
 cboPord.SetFocus
 cmdNew.Enabled = False
Else
'put all stock in Stock In table
 rsStockin.Open "select * from Stock_In", con, adOpenKeyset, adLockOptimistic
  For i = 1 To grdStockIn.Rows - 1
     rsStockin.AddNew
     rsStockin.Fields(0) = lblStockIn.Caption
     rsStockin.Fields(1) = cboPord.Text
     rsStockin.Fields(2) = Date
     rsStockin.Fields(3) = grdStockIn.TextMatrix(i, 0)
     rsStockin.Fields(4) = grdStockIn.TextMatrix(i, 4)
     rsStockin.Fields(5) = txtSup.Text
     rsStockin.Update
     'add the unit in price to the IMEI table
     'rs2.Open "select * from IMEI where ModelNo='" & grdStockIn.TextMatrix(i, 0) & "'", con, adOpenKeyset, adLockOptimistic
     'If rs2.RecordCount > 0 Then
     'rs2.Fields![ModelNo] = Val(grdStockIn.TextMatrix(i, 0))
     'rs2.Fields![Unit_Price] = Val(grdStockIn.TextMatrix(i, 2))
     'rs2.Update
    'End If
    
       'Set rs2 = Nothing
      
Next
    Set rsStockin = Nothing
    
    'delete PO that already stock in
    'rs.Open "delete from PordNo where PordNo='" & cboPord.Text & "'", con, adOpenKeyset, adLockOptimistic
    'rs.Close
    rsStockin.Open "delete from P_Order where PO_No='" & cboPord.Text & "'", con, adOpenKeyset, adLockOptimistic
    MsgBox "Stock successfully added", vbOKOnly + vbInformation, "Success"
    cmdNew.Caption = "New"
   
    'clear and disable the components
    cboPord.Clear
    cboPord.Enabled = False
    grdStockIn.Enabled = False
    cboPord.Text = ""
    txtSupName.Text = ""
    txtDateOrd.Text = ""
    txtDaterec.Text = ""
    grdStockIn.Rows = 1
    lblStockIn.Caption = ""
    Set rsStockin = Nothing
    
 End If
 'cmdNew.Enabled = False
 
End Sub

Private Sub Form_Load()
frmStockIn.Top = 350
frmStockIn.Left = 4100

grdStockIn.TextMatrix(0, 0) = "ModelNo"
grdStockIn.TextMatrix(0, 1) = "ModelName"
grdStockIn.TextMatrix(0, 2) = "Price"
grdStockIn.TextMatrix(0, 3) = "Order"
grdStockIn.TextMatrix(0, 4) = "Recieve"

With grdStockIn
   .ColWidth(0) = 1300
   .ColWidth(1) = 3300
   .ColWidth(2) = 1100
   .ColWidth(3) = 1100
   .ColWidth(4) = 1100
End With
Call Connect


End Sub

Sub createID()
Dim lastno As Long, newno As Long
Set rsStockin = New ADODB.Recordset
rsStockin.Open "select * from Stock_In", con, adOpenDynamic, adLockOptimistic

With rsStockin
 If .BOF = True And .EOF = True Then
 lastno = 0
 Else
 .MoveLast
 lastno = CLng(Mid(.Fields(0), 3, 2))
 End If
 
  newno = lastno + 1
   lblStockIn.Caption = "SI" & newno
End With
Set rsStockin = Nothing

End Sub

Private Sub grdStockIn_Click()
frmStockRec.txtModelNo = grdStockIn.TextMatrix(grdStockIn.Row, 0)
frmStockRec.lblProdName.Caption = grdStockIn.TextMatrix(grdStockIn.Row, 1)
frmStockRec.Show vbModal

End Sub

