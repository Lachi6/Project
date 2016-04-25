VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchaseReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPurchaseReport.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   7740
   Begin VB.CommandButton cmdSup 
      Caption         =   "By Supplier"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      ToolTipText     =   "Report Of Purchase Product On the Basis Of supplier"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdSelected 
      Caption         =   "View Report"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "View Report Between Date"
      Top             =   1440
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPickerTo 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102694913
      CurrentDate     =   42260
   End
   Begin MSComCtl2.DTPicker DTPickerFrom 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102694913
      CurrentDate     =   42260
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Whole Report"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Whole Report Till Date"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Report According Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   240
      Top             =   720
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "frmPurchaseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAll_Click()
Call Connect
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strsql As String
strsql = "select * from Stock_In"
    With rs
   If rs.State = 1 Then .Close
    .Open strsql, con, adOpenDynamic, adLockBatchOptimistic
    If .RecordCount <> 0 Then
        
    Set PurchaseReport.DataSource = rs
    PurchaseReport.Show
        Else
            MsgBox "No Record Available!", vbInformation, "Ambica"
           End If
           End With
End Sub

Private Sub cmdSelected_Click()
Call Connect
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strsql As String
'select * from Sales where Sale_Date BETWEEN #" & DTPickerFrom.Value & "# And #" & DTPickerTo.Value & "#"
strsql = "select * from Stock_In where Date_Recieved BETWEEN #" & DTPickerFrom.Value & "# And #" & DTPickerTo.Value & "#"

    With rs
   If rs.State = 1 Then .Close
    .Open strsql, con, adOpenDynamic, adLockBatchOptimistic
    If .RecordCount <> 0 Then
            
        Set PurchaseReport.DataSource = rs
        PurchaseReport.Show
        PurchaseReport.Refresh
        
    Else
            MsgBox "No Record Available!", vbInformation, "Ambica"
           End If
           End With

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSup_Click()
Call Connect

On Error Resume Next
Dim RPT, RPT2
RPT = InputBox("Enter product supplier name.", , "Enter here")

If RPT = "" Then
MsgBox "Enter Supplier Name", vbOKOnly + vbInformation, "Warning"

Else
Set rsRpt = New ADODB.Recordset
rsRpt.Open "SELECT * From Stock_In where Supplier_Name='" & RPT & "'", con


RPT2 = rsRpt!Supplier_name
Set PBS.DataSource = rsRpt.DataSource

For Each obj In PBS.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rsRpt.DataMember
    End If
Next
PBS.Sections("Section1").Controls("Text1").DataField = "SIn_no"
PBS.Sections("Section1").Controls("Text2").DataField = "PO_No"
PBS.Sections("Section2").Controls("Label1").Caption = RPT2
PBS.Sections("Section1").Controls("Text4").DataField = "Date_Recieved"
PBS.Sections("Section1").Controls("Text5").DataField = "ModelNo"
PBS.Sections("Section1").Controls("Text6").DataField = "Quantity"
PBS.Refresh
PBS.Show
Set rsRpt = Nothing

End If
con.Close
End Sub

Private Sub Form_Load()
frmPurchaseReport.Top = 350
frmPurchaseReport.Left = 4100
'using 2 dtpickers
DTPickerFrom.Value = Date
DTPickerTo.Value = Date

End Sub


