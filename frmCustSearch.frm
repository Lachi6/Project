VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Customer Details"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCustSearch.frx":0000
   ScaleHeight     =   3435
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Height          =   615
      Left            =   8280
      Picture         =   "frmCustSearch.frx":9CAA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close"
      Top             =   840
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid info 
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1508
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBillNo 
      Caption         =   "By BillNo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      ToolTipText     =   "Search Customer usong BillNo"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "By Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Search Customer By Name"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdPhno 
      Appearance      =   0  'Flat
      Caption         =   "By Phone No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Search Customer By Phone Number"
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtSearch 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Enter Your Text Here"
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Left            =   480
      Top             =   720
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Customer's Detail"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmCustSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdBillNo_Click()
cmdPrint.Enabled = True
Set rs = New ADODB.Recordset
If rs.State = 1 Then
rs.Close
End If
rs.CursorLocation = adUseClient
    rs.Open "select * from Sales where SalesNo='" & txtSearch.Text & "'", con, adOpenDynamic, adLockOptimistic
    Set info.DataSource = rs
    If rs.RecordCount = 0 Then
        MsgBox ("Record Not Found")
        Exit Sub
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdName_Click()
Set rs = New ADODB.Recordset
If rs.State = 1 Then
rs.Close
End If
rs.CursorLocation = adUseClient
    rs.Open "select * from Sales where CustomerName='" & txtSearch.Text & "'", con, adOpenDynamic, adLockOptimistic
    Set info.DataSource = rs
    If rs.RecordCount = 0 Then
        MsgBox ("Record Not Found")
        Exit Sub
    End If
End Sub

Private Sub cmdPhno_Click()
Set rs = New ADODB.Recordset
If rs.State = 1 Then
rs.Close
End If
rs.CursorLocation = adUseClient
    rs.Open "select * from Sales where CustomerPhone='" & txtSearch.Text & "'", con, adOpenDynamic, adLockOptimistic
    Set info.DataSource = rs
    If rs.RecordCount = 0 Then
        MsgBox ("Record Not Found")
        Exit Sub
    End If
End Sub
Private Sub cmdPrint_Click()
'If DataEnvironment1.rs.State = 1 Then
 '     DataEnvironment1.rs.Close
'End If
  
 '   DataEnvironment1.Commands.Item("cmdsale").Parameters.Item(0).Value = txtsale
    
  '  SaleBill.Show

Set rsRpt = New ADODB.Recordset
rsRpt.Open "SELECT * From Sales , Model where Model.ModelNo=Sales.ModelNo AND SalesNo like '" & txtSearch.Text & "'", con
Set salesBill.DataSource = rsRpt.DataSource
'salesBill.Refresh
salesBill.Show
Set rsRpt = Nothing
End Sub
Private Sub Form_Load()
frmCustSearch.Top = 350
frmCustSearch.Left = 4100

End Sub

