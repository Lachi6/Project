VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPurchaseOrd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order"
   ClientHeight    =   8460
   ClientLeft      =   4215
   ClientTop       =   795
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmPurchaseOrd.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Picture         =   "frmPurchaseOrd.frx":9CAA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7680
      Width           =   1215
   End
   Begin VB.ComboBox Combosup 
      Height          =   315
      Left            =   2520
      TabIndex        =   21
      Top             =   1920
      Width           =   3255
   End
   Begin VB.ComboBox cboBrand 
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
      Height          =   360
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtPOrdNo 
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
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H8000000E&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Picture         =   "frmPurchaseOrd.frx":A12A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H8000000E&
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Picture         =   "frmPurchaseOrd.frx":A46C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H8000000E&
      Caption         =   "&Purchase"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Picture         =   "frmPurchaseOrd.frx":A7AE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000E&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Picture         =   "frmPurchaseOrd.frx":AAF0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpDateReq 
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   -2147483624
      Format          =   102629377
      CurrentDate     =   42075
   End
   Begin MSComCtl2.DTPicker dtpDateOrd 
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   -2147483624
      Format          =   102629377
      CurrentDate     =   42075
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1455
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   -2147483624
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ListView lvwOrder 
      Height          =   1575
      Left            =   360
      TabIndex        =   9
      Top             =   5880
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2778
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ModelNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Model Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Brand"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name"
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
      Left            =   360
      TabIndex        =   20
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
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
      Left            =   3120
      TabIndex        =   19
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblDateReq 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Required"
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
      Left            =   6240
      TabIndex        =   18
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblDateOrd 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Order"
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
      Left            =   6240
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblSupplier 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand's Name"
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
      TabIndex        =   16
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblOrdNo 
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
      Left            =   360
      TabIndex        =   15
      Top             =   960
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   1815
      Left            =   240
      Top             =   600
      Width           =   9615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   600
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      Height          =   2055
      Left            =   240
      Top             =   2760
      Width           =   9615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Product List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   2055
      Left            =   240
      Top             =   5520
      Width           =   9615
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   2880
      Picture         =   "frmPurchaseOrd.frx":AE32
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Order List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name :"
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
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblItemName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   4920
      Width           =   3615
   End
End
Attribute VB_Name = "frmPurchaseOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsso As New ADODB.Recordset
Dim rssocid As New ADODB.Recordset

Private Sub cboBrand_click()

rsso.Open "select * from Model where BrandName='" & cboBrand.Text & "'", con, adOpenKeyset, adLockPessimistic
If Not (rsso.BOF And rsso.EOF) Then MSHFlexGrid1.Enabled = True
Set MSHFlexGrid1.DataSource = rsso
Set rsso = Nothing
End Sub



Private Sub cmdAdd_Click()
If Combosup.Text = "" Then
MsgBox ("Please Select Supplier")
Else
With frmAddQuanP
   .lblProdID.Caption = scAd
   .lblProdName.Caption = itmAd
End With
frmAddQuanP.Show vbModal

cmdPurchase.Enabled = True
cmdAdd.Enabled = False
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Call Connect
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strsql As String
strsql = "select * from P_Order where PO_No='" & txtPOrdNo.Text & "'"
    With rs
   If rs.State = 1 Then .Close
    .Open strsql, con, adOpenDynamic, adLockBatchOptimistic
    If .RecordCount <> 0 Then
        
    Set PurchaseOrderReport.DataSource = rs
    PurchaseOrderReport.Show
        Else
            MsgBox "No Record Available!", vbInformation, "Ambica"
           End If
           End With
End Sub

Private Sub cmdPurchase_Click()
Dim cust, DOr, DRe, SC, QTY, SoN, sup As String
SoN = txtPOrdNo.Text
'cust = cboBrand.Text
DOr = dtpDateOrd.Value
DRe = dtpDateReq.Value
sup = Combosup.Text


For i = 1 To lvwOrder.ListItems.Count
SC = lvwOrder.ListItems(i).Text
QTY = lvwOrder.ListItems(i).ListSubItems(2).Text
cust = lvwOrder.ListItems(i).ListSubItems(3).Text

rsso.Open "Select * from P_Order", con, adOpenDynamic, adLockOptimistic
With rsso
  .AddNew
  .Fields(0) = SoN
  .Fields(1) = SC
  .Fields(2) = QTY
  .Fields(3) = cust
  .Fields(4) = sup
  .Fields(5) = DOr
  .Fields(6) = DRe
  .Update
End With
rsso.Close

Next

lvwOrder.ListItems.Clear
cmdPurchase.Enabled = False
cmdAdd.Enabled = False
MsgBox "Purchase transaction no " & SoN & " has been purchased. ", vbOKOnly + vbInformation, "Success"
Set rsso = Nothing

Set rssocid = New ADODB.Recordset
rssocid.Open "select * from PordNo", con, adOpenDynamic, adLockOptimistic
With rssocid
   '.MoveLast
   .AddNew
   .Fields(0) = SoN
   .Update
End With
Set rssocid = Nothing

cmdPrint.Enabled = True
End Sub

Private Sub cmdRemove_Click()
Dim i As Integer
For i = lvwOrder.ListItems.Count To 1 Step -1
   If lvwOrder.ListItems(i).Checked = True Then
     lvwOrder.ListItems.Remove i
   Else
   MsgBox "Please select an item from Order List.", vbOKOnly + vbInformation, "Warning"
   Exit Sub
   End If
 Next i
 
End Sub



Private Sub Combosup_Click()
Combosup.locked = True
End Sub

Private Sub dtpDateReq_Change()
If dtpDateReq.Value < Date Then
MsgBox ("Date Required Cannot Be Less Than Current Date")
dtpDateReq.Value = Date
dtpDateReq.SetFocus
End If
End Sub



Private Sub Form_Load()
cmdAdd.Enabled = False
frmPurchaseOrd.Top = 350
frmPurchaseOrd.Left = 4100

dtpDateOrd.Value = Date
dtpDateReq.Value = Date
Call Connect
Set rsso = New ADODB.Recordset
rsso.Open "select * from P_Order", con, adOpenDynamic, adLockOptimistic
With MSHFlexGrid1
 .ColWidth(0) = 300
 .ColWidth(1) = 1500
 .ColWidth(2) = 3000
End With
Call createID
Set rsso = Nothing
txtPOrdNo.locked = True

rsso.Open "select * from Brand", con, adOpenKeyset, adLockPessimistic
If rsso.RecordCount > 0 Then
Do Until rsso.EOF
  cboBrand.AddItem rsso.Fields(1).Value
  rsso.MoveNext
Loop
End If
Set rsso = Nothing

rsso.Open "select * from Suppliers", con, adOpenKeyset, adLockPessimistic
If rsso.RecordCount > 0 Then
Do Until rsso.EOF
  Combosup.AddItem rsso.Fields(1).Value
  rsso.MoveNext
Loop
End If
Set rsso = Nothing


End Sub
Sub createID()
Dim lastno As Long, newno As Long
Set rssocid = New ADODB.Recordset
rssocid.Open "select * from PordNo", con, adOpenDynamic, adLockOptimistic
With rssocid
 If .BOF = True And .EOF = True Then
 lastno = 0
 Else
 .MoveLast
 lastno = CLng(Mid(.Fields(0), 3, 2))
 End If
 
  newno = lastno + 1
  txtPOrdNo.Text = "PO" & newno
End With
Set rssocid = Nothing
End Sub




Private Sub MSHFlexGrid1_Click()
i = MSHFlexGrid1.Row
With MSHFlexGrid1
   lblItemName.Caption = .TextMatrix(i, 2)
   scAd = .TextMatrix(i, 1)
   itmAd = .TextMatrix(i, 2)
End With
If Not scAd = "" Then
   cmdAdd.Enabled = True
   cmdAdd.SetFocus
End If


End Sub


