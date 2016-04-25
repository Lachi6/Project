VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSuppliers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier"
   ClientHeight    =   8490
   ClientLeft      =   4035
   ClientTop       =   3600
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmSuppliers.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   4920
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102694913
      CurrentDate     =   42260
   End
   Begin VB.TextBox txtSupID 
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtSupName 
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
      Left            =   3360
      TabIndex        =   13
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txtSupMail 
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
      Left            =   3360
      TabIndex        =   12
      Top             =   4200
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
      Left            =   6360
      Picture         =   "frmSuppliers.frx":9CAA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Go to last record"
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
      Left            =   5040
      Picture         =   "frmSuppliers.frx":A8EC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Go to previous record"
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
      Left            =   2760
      Picture         =   "frmSuppliers.frx":B52E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Go to next record"
      Top             =   6000
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
      Left            =   1440
      Picture         =   "frmSuppliers.frx":C170
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Go to first record"
      Top             =   6000
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
      Left            =   8520
      Picture         =   "frmSuppliers.frx":CDB2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Update current record"
      Top             =   5160
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
      Left            =   8520
      Picture         =   "frmSuppliers.frx":D9F4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add new Products"
      Top             =   1080
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
      Left            =   8520
      Picture         =   "frmSuppliers.frx":E636
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
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
      Left            =   8520
      Picture         =   "frmSuppliers.frx":F278
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "close this window"
      Top             =   6120
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
      Left            =   8520
      Picture         =   "frmSuppliers.frx":FEBA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Edit current item"
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000E&
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
      Left            =   8520
      Picture         =   "frmSuppliers.frx":10AFC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save new item"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtSupAddr 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox txtSupPhone 
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
      Left            =   3360
      TabIndex        =   0
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers"
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
      Left            =   3600
      TabIndex        =   22
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3480
      Picture         =   "frmSuppliers.frx":1173E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblSupDate 
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
      Left            =   1680
      TabIndex        =   21
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   4575
      Left            =   1080
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label lblSupID 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID"
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
      TabIndex        =   20
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblSupName 
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
      Left            =   1680
      TabIndex        =   19
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label LblSupAddr 
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
      Left            =   1680
      TabIndex        =   18
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblSupPhone 
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
      Left            =   1320
      TabIndex        =   17
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblSupMail 
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
      Left            =   1680
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers Details"
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
      Left            =   3120
      TabIndex        =   15
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      Height          =   1335
      Left            =   1080
      Top             =   5760
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   8280
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      Height          =   3015
      Left            =   8280
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssupplier As New ADODB.Recordset
Dim rsaddsup As New ADODB.Recordset

Private Sub cmdAdd_Click()
enabletxt
cmdClear.Enabled = True
cmdSave.Enabled = True
Clear

Call createID
DTPicker1.Value = Date
Set rsaddsup = New ADODB.Recordset
rsaddsup.Open "select * from Suppliers", con, adOpenDynamic, adLockOptimistic
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
rssupplier.MoveFirst
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
With rssupplier
.MoveFirst

End With
cmdCancel.Enabled = True
End Sub

Private Sub cmdFirst_Click()
enablebutton

On Error Resume Next
rssupplier.MoveFirst
  If rssupplier.BOF Then
      MsgBox "You are at first record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdLast_Click()
enablebutton

On Error Resume Next
rssupplier.MoveLast
  If rssupplier.EOF Then
      MsgBox "You are at last record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdNext_Click()
enablebutton

If rssupplier.BOF = True And rssupplier.EOF = True Then Exit Sub
On Error Resume Next
rssupplier.MoveNext
  If rssupplier.EOF Then
      MsgBox "You are at last record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdPrev_Click()
enablebutton

If rssupplier.BOF = True And rssupplier.EOF = True Then Exit Sub
On Error Resume Next
rssupplier.MovePrevious
  If rssupplier.BOF Then
      MsgBox "You are at first record", vbOKOnly + vbInformation, "Warning"
  End If

showdata
End Sub

Private Sub cmdSave_Click()


 If txtSupID.Text = "" Then
 MsgBox "Please enter Supplier ID.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtSupName.Text = "" Then
 MsgBox "plaese enter Supplier name.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtSupAddr.Text = "" Then
 MsgBox "plaese enter supplier's address.", vbOKOnly + vbInformation, "Error"

 ElseIf txtSupPhone.Text = "" Then
 MsgBox "plaese enter supplier's Mobile number.", vbOKOnly + vbInformation, "Error"

 ElseIf txtSupMail.Text = "" Then
 MsgBox "plaese enter supplier's email id.", vbOKOnly + vbInformation, "Error"
 
ElseIf DTPicker1.Value = "" Then
 MsgBox "Please enter date.", vbOKOnly + vbInformation, "Error"
 
 Else
   With rssupplier
   .AddNew
   .Fields(0) = txtSupID.Text
   .Fields(1) = txtSupName.Text
   .Fields(2) = txtSupAddr.Text
   .Fields(3) = txtSupPhone.Text
   .Fields(4) = txtSupMail.Text
   .Fields(5) = txtdate.Text
   .Update
   
   MsgBox "Supplier Successfully Added", vbOKOnly, "Success"
   End With
   
   Clear
   End If
End Sub

Private Sub cmdUpdate_Click()
Dim rsupdate As New ADODB.Recordset
Set rsupdate = New ADODB.Recordset
rsupdate.Open "select * from Suppliers where Supplier_ID='" & txtSupID.Text & "'", con, adOpenDynamic, adLockOptimistic

If txtSupID.Text = "" Then
 MsgBox "Please enter Supplier ID.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtSupName.Text = "" Then
 MsgBox "plaese enter Supplier name.", vbOKOnly + vbInformation, "Error"
 
 ElseIf txtSupAddr.Text = "" Then
 MsgBox "plaese enter Supplier's address.", vbOKOnly + vbInformation, "Error"

 ElseIf txtSupPhone.Text = "" Then
 MsgBox "plaese enter Supplier's Mobile number.", vbOKOnly + vbInformation, "Error"

 ElseIf txtSupMail.Text = "" Then
 MsgBox "plaese enter Supplier's email id.", vbOKOnly + vbInformation, "Error"
 
 ElseIf DTPicker1.Value = "" Then
 MsgBox "Please enter date.", vbOKOnly + vbInformation, "Error"
 
  Else
  With rsupdate
   .Fields(0) = txtSupID.Text
   .Fields(1) = txtSupName.Text
   .Fields(2) = txtSupAddr.Text
   .Fields(3) = txtSupPhone.Text
   .Fields(4) = txtSupMail.Text
   .Fields(5) = DTPicker1.Value
  End With
  rsupdate.Update
  MsgBox "Record updated successfully", vbOKOnly, "Success"
  Clear
  disabletxt
  End If
  
End Sub

Private Sub Form_Load()
disablecontrol
frmSuppliers.Top = 320
frmSuppliers.Left = 4100
Call Connect
Set rssupplier = New ADODB.Recordset
rssupplier.Open "select * from Suppliers", con, adOpenDynamic, adLockOptimistic
End Sub
Sub showdata()

txtSupID.Text = rssupplier.Fields(0)
txtSupName.Text = rssupplier.Fields(1)
txtSupAddr.Text = rssupplier.Fields(2)
txtSupPhone.Text = rssupplier.Fields(3)
txtSupMail.Text = rssupplier.Fields(4)
DTPicker1.Value = rssupplier.Fields(5)
End Sub
Sub disablecontrol()

cmdCancel.Enabled = False
cmdEdit.Enabled = False
cmdUpdate.Enabled = False
cmdSave.Enabled = False
cmdClear.Enabled = False

txtSupID.locked = True
txtSupName.locked = True
txtSupAddr.locked = True
txtSupPhone.locked = True
txtSupMail.locked = True
'DTPicker1.locked = True

End Sub

Sub enablebutton()
cmdAdd.Enabled = True
cmdEdit.Enabled = True

If cmdCancel.Enabled = True Then
cmdCancel.Enabled = False
End If

End Sub
Sub enabletxt()

txtSupName.locked = False
txtSupAddr.locked = False
txtSupPhone.locked = False
txtSupMail.locked = False

End Sub
Sub disabletxt()

txtSupID.locked = True
txtSupName.locked = True
txtSupAddr.locked = True
txtSupPhone.locked = True
txtSupMail.locked = True
'txtdate.locked = True
End Sub
Sub Clear()
txtSupName.Text = ""
txtSupAddr.Text = ""
txtSupPhone.Text = ""
txtSupMail.Text = ""

End Sub
Sub createID()
Dim lastno As Long, newno As Long
'check if there are record in the file
With rssupplier
 If .BOF = True And .EOF = True Then
 lastno = 1
 Else
 .MoveLast
 lastno = CLng(Mid(.Fields(0), 2, 2))
 End If
 'generate new no
  newno = lastno + 1
  txtSupID.Text = "S" & newno
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
  
Private Sub txtSupName_KeyPress(KeyAscii As Integer)
Call DonotAllowSpChar(KeyAscii)
End Sub


Private Sub txtSupPhone_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
End Sub

