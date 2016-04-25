VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmProduct 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Products"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmProduct.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdModel 
      Caption         =   "Add Model"
      Height          =   735
      Left            =   9360
      Picture         =   "frmProduct.frx":E0F9
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdBrand 
      Caption         =   "Add Brand"
      Height          =   735
      Left            =   9360
      Picture         =   "frmProduct.frx":E594
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   735
      Left            =   5400
      Picture         =   "frmProduct.frx":EA2F
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   3480
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   1455
      Left            =   2280
      TabIndex        =   3
      Top             =   3960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2566
      _Version        =   393216
      GridColor       =   12582912
      GridColorFixed  =   8421376
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox ComboModel 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ComboBox ComboBrand 
      Height          =   315
      Left            =   6600
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1455
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   -2147483624
      GridColor       =   65280
      GridColorFixed  =   8454016
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Model-No To View Model Price And IMEI Number And Total Quantity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Brand To View All The Models Avaliable"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Details Of Product"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   360
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      Height          =   2175
      Left            =   1920
      Top             =   3360
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   1920
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Model"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Brand"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsso As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Private Sub cmdBrand_Click()
frmBrand.Show
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdModel_Click()
frmModel.Show
End Sub

Private Sub cmdok_Click()
Unload Me
End Sub

Private Sub ComboBrand_Click()
ComboBrand.locked = True
Set rsso = New ADODB.Recordset
MSHFlexGrid1.Visible = True
rsso.Open "select * from Model where BrandName='" & ComboBrand.Text & "'", con, adOpenKeyset, adLockPessimistic
If Not (rsso.BOF And rsso.EOF) Then MSHFlexGrid1.Enabled = True
Set MSHFlexGrid1.DataSource = rsso
Set rsso = Nothing
ComboModel.Visible = True
Label5.Visible = True
Label2.Visible = True
Shape2.Visible = True


Set rs = New ADODB.Recordset
rs.Open "select * from Model where BrandName='" & ComboBrand.Text & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
Do Until rs.EOF
  ComboModel.AddItem rs.Fields(0).Value
  rs.MoveNext
Loop
End If
Set rs = Nothing

End Sub

Private Sub ComboModel_Click()
'ComboModel.locked = True
Set rsso = New ADODB.Recordset

rsso.Open "select * from IMEI where ModelNo='" & ComboModel.Text & "'", con, adOpenKeyset, adLockPessimistic
If Not (rsso.BOF And rsso.EOF) Then MSHFlexGrid2.Enabled = True
Set MSHFlexGrid2.DataSource = rsso
txtQty.Text = rsso.RecordCount
Set rsso = Nothing
MSHFlexGrid2.Visible = True
txtQty.Visible = True
Label6.Visible = True
End Sub

Private Sub Form_Load()

frmProduct.Top = 350
frmProduct.Left = 4100
MSHFlexGrid2.Visible = False
MSHFlexGrid1.Visible = False
ComboModel.Visible = False
Label5.Visible = False
Label2.Visible = False
Shape2.Visible = False
Label6.Visible = False
txtQty.Visible = False

'Call Connect
If rsso.State = 1 Then
rsso.Close
End If
Set rsso = New ADODB.Recordset
rsso.Open "select * from Brand", con, adOpenKeyset, adLockPessimistic
If rsso.RecordCount > 0 Then
Do Until rsso.EOF
  ComboBrand.AddItem rsso.Fields(1).Value
  rsso.MoveNext
Loop
End If

Set rsso = Nothing
With MSHFlexGrid1
 .ColWidth(0) = 300
 .ColWidth(1) = 1500
 .ColWidth(2) = 3000
End With
With MSHFlexGrid2
 .ColWidth(0) = 300
 .ColWidth(1) = 1500
 .ColWidth(2) = 3000
End With

End Sub

'Private Sub Timer1_Timer()
'Label3.Top = frmProduct.Height / 2
'Label3.Left = Label3.Left - 30
'If Label3.Left < 0 - Label3.Width Then
'Label3.Left = frmProduct.Width
'End If
'End Sub

Sub Clear()
ComboBrand.Text = ""
ComboModel.Text = ""
End Sub

