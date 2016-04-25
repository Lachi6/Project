VERSION 5.00
Begin VB.Form frmCustomers 
   Caption         =   "Customer"
   ClientHeight    =   8610
   ClientLeft      =   4110
   ClientTop       =   3675
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   Picture         =   "frmCustomers.frx":0000
   ScaleHeight     =   8610
   ScaleWidth      =   11175
   Begin VB.TextBox txtCustPhone 
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
      Left            =   3000
      TabIndex        =   23
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtCustAddr 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   3000
      Width           =   3855
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000009&
      Caption         =   "&Save"
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
      Left            =   8160
      Picture         =   "frmCustomers.frx":9CAA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Save new item"
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000E&
      Caption         =   "&Edit"
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
      Left            =   8160
      Picture         =   "frmCustomers.frx":A8EC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Edit current item"
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000E&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Enabled         =   0   'False
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
      Left            =   8160
      Picture         =   "frmCustomers.frx":B52E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "close this window"
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H8000000E&
      Caption         =   "&Clear"
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
      Left            =   8160
      Picture         =   "frmCustomers.frx":C170
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H80000014&
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
      Height          =   855
      Left            =   8160
      Picture         =   "frmCustomers.frx":CDB2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add new Products"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H8000000E&
      Caption         =   "&Update"
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
      Left            =   8160
      Picture         =   "frmCustomers.frx":D9F4
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Update current record"
      Top             =   5160
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
      Left            =   1080
      Picture         =   "frmCustomers.frx":E636
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Go to first record"
      Top             =   6000
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
      Left            =   2400
      Picture         =   "frmCustomers.frx":F278
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Go to next record"
      Top             =   6000
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
      Left            =   4680
      Picture         =   "frmCustomers.frx":FEBA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Go to previous record"
      Top             =   6000
      Width           =   975
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
      Left            =   6000
      Picture         =   "frmCustomers.frx":10AFC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Go to last record"
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox txtCustMail 
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
      Left            =   3000
      TabIndex        =   3
      Top             =   4200
      Width           =   3855
   End
   Begin VB.TextBox txtCustName 
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
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txtCustID 
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1800
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
      Left            =   3000
      TabIndex        =   0
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customers"
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
      Left            =   3840
      TabIndex        =   21
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      Height          =   3015
      Left            =   7920
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   7920
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      Height          =   1335
      Left            =   720
      Top             =   5760
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customers Details"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3720
      Picture         =   "frmCustomers.frx":1173E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblCustMail 
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
      Left            =   1320
      TabIndex        =   19
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblCustPhone 
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
      Left            =   960
      TabIndex        =   18
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label LblCustAddr 
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
      Left            =   1320
      TabIndex        =   17
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblCustName 
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
      Left            =   1320
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblCustID 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   4575
      Left            =   720
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label lblCustDate 
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
      Left            =   1320
      TabIndex        =   14
      Top             =   4920
      Width           =   1575
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rscustomer As New ADODB.Recordset
Dim rsaddsup As New ADODB.Recordset

Private Sub cmdAdd_Click()
enabletxt
cmdClear.Enabled = True
cmdSave.Enabled = True
Clear

Call createID
txtdate.Text = Date
Set rsaddsup = New ADODB.Recordset
rsaddsup.Open "select * from Customers", con, adOpenDynamic, adLockOptimistic
With rsaddsup
.MoveFirst

End With
navigatedisable
cmdCancel.Enabled = True
cmdEdit.Enabled = False
cmdUpdate.Enabled = False

End Sub

Private Sub cmdCancel_Click()

navigateenable
cmdClear.Enabled = False
cmdSave.Enabled = False
cmdEdit.Enabled = False
cmdUpdate.Enabled = False
If cmdAdd.Enabled = False Then
cmdAdd.Enabled = True
End If
disabletxt
rscustomer.MoveFirst
showdata
End Sub

Private Sub cmdClear_Click()
enabletxt
Clear
End Sub

Private Sub cmdEdit_Click()
navigatedisable
enabletxt
cmdAdd.Enabled = False
cmdClear.Enabled = False
cmdSave.Enabled = False

cmdUpdate.Enabled = True
With rscustomer
.MoveFirst

End With
cmdCancel.Enabled = True
End Sub

Private Sub cmdFirst_Click()
enablebutton

On Error Resume Next
rscustomer.MoveFirst
  If rscustomer.BOF Then
      MsgBox "You are at first record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdLast_Click()
enablebutton

On Error Resume Next
rscustomer.MoveLast
  If rscustomer.EOF Then
      MsgBox "You are at last record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdNext_Click()
enablebutton

If rscustomer.BOF = True And rscustomer.EOF = True Then Exit Sub
On Error Resume Next
rscustomer.MoveNext
  If rscustomer.EOF Then
      MsgBox "You are at last record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdPrev_Click()
enablebutton

If rscustomer.BOF = True And rscustomer.EOF = True Then Exit Sub
On Error Resume Next
rscustomer.MovePrevious
  If rscustomer.BOF Then
      MsgBox "You are at first record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdSave_Click()


 If txtCustID.Text = "" Then
 MsgBox "Please enter Customer ID.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtCustName.Text = "" Then
 MsgBox "plaese enter customer name.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtCustAddr.Text = "" Then
 MsgBox "plaese enter customer's address.", vbOKOnly + vbInformation, "Error"

 ElseIf txtCustPhone.Text = "" Then
 MsgBox "plaese enter customer's Mobile number.", vbOKOnly + vbInformation, "Error"

 ElseIf txtCustMail.Text = "" Then
 MsgBox "plaese enter customer's email id.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtdate.Text = "" Then
 MsgBox "Please enter date.", vbOKOnly + vbInformation, "Error"
 
 Else
   With rscustomer
   .AddNew
   .Fields(0) = txtCustID.Text
   .Fields(1) = txtCustName.Text
   .Fields(2) = txtCustAddr.Text
   .Fields(3) = txtCustPhone.Text
   .Fields(4) = txtCustMail.Text
   .Fields(5) = txtdate.Text
   .Update
   
   MsgBox "Customer Successfully Added", vbOKOnly, "Success"
   End With
   
   Clear
   End If
End Sub

Private Sub cmdUpdate_Click()
Dim rsupdate As New ADODB.Recordset
Set rsupdate = New ADODB.Recordset
rsupdate.Open "select * from Customers where Customer_ID='" & txtCustID.Text & "'", con, adOpenDynamic, adLockOptimistic

If txtCustID.Text = "" Then
 MsgBox "Please enter Customer ID.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtCustName.Text = "" Then
 MsgBox "plaese enter customer name.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtCustAddr.Text = "" Then
 MsgBox "plaese enter customer's address.", vbOKOnly + vbInformation, "Error"

 ElseIf txtCustPhone.Text = "" Then
 MsgBox "plaese enter customer's Mobile number.", vbOKOnly + vbInformation, "Error"

 ElseIf txtCustMail.Text = "" Then
 MsgBox "plaese enter customers email id.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtdate.Text Then
 MsgBox "Please enter date.", vbOKOnly + vbInformation, "Error"
 
  Else
  With rsupdate
   .Fields(0) = txtCustID.Text
   .Fields(1) = txtCustName.Text
   .Fields(2) = txtCustAddr.Text
   .Fields(3) = txtCustPhone.Text
   .Fields(4) = txtCustMail.Text
   .Fields(5) = txtdate.Text
  End With
  rsupdate.Update
  MsgBox "Record updated successfully", vbOKOnly, "Success"
  Clear
  disabletxt
  End If
  
End Sub

Private Sub Form_Load()
frmCustomers.Top = 2000
frmCustomers.Left = 4100
disablecontrol

Call Connect
Set rscustomer = New ADODB.Recordset
rscustomer.Open "select * from Customers", con, adOpenDynamic, adLockOptimistic
End Sub
Sub showdata()

txtCustID.Text = rscustomer.Fields(0)
txtCustName.Text = rscustomer.Fields(1)
txtCustAddr.Text = rscustomer.Fields(2)
txtCustPhone.Text = rscustomer.Fields(3)
txtCustMail.Text = rscustomer.Fields(4)
txtdate.Text = rscustomer.Fields(5)
End Sub
Sub disablecontrol()

cmdCancel.Enabled = False
cmdEdit.Enabled = False
cmdUpdate.Enabled = False
cmdSave.Enabled = False
cmdClear.Enabled = False

txtCustID.locked = True
txtCustName.locked = True
txtCustAddr.locked = True
txtCustPhone.locked = True
txtCustMail.locked = True
txtdate.locked = True

End Sub

Sub enablebutton()
cmdAdd.Enabled = True
cmdEdit.Enabled = True

If cmdCancel.Enabled = True Then
cmdCancel.Enabled = False
End If

End Sub
Sub enabletxt()

txtCustName.locked = False
txtCustAddr.locked = False
txtCustPhone.locked = False
txtCustMail.locked = False

End Sub
Sub disabletxt()

txtCustID.locked = True
txtCustName.locked = True
txtCustAddr.locked = True
txtCustPhone.locked = True
txtCustMail.locked = True
txtdate.locked = True
End Sub
Sub Clear()
txtCustName.Text = ""
txtCustAddr.Text = ""
txtCustPhone.Text = ""
txtCustMail.Text = ""

End Sub
Sub createID()
Dim lastno As Long, newno As Long
'check if there are record in the file
With rscustomer
 If .BOF = True And .EOF = True Then
 lastno = 1
 Else
 .MoveLast
 lastno = CLng(Mid(.Fields(0), 2, 2))
 End If
 'generate new no
  newno = lastno + 1
  txtCustID.Text = "C" & newno
End With
End Sub

Sub navigateenable()
'enable all navigation control
cmdFirst.Enabled = True
cmdLast.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
End Sub

Sub navigatedisable()
'disable all navigation control
cmdFirst.Enabled = False
cmdLast.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
End Sub
  
Private Sub txtCustName_KeyPress(KeyAscii As Integer)
Call DonotAllowSpChar(KeyAscii)
End Sub


Private Sub txtCustPhone_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
End Sub
