VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSaleReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmSaleReport.frx":0000
   ScaleHeight     =   4830
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelected 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Show Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPickerTo 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102694913
      CurrentDate     =   42246
   End
   Begin MSComCtl2.DTPicker DTPickerFrom 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102694913
      CurrentDate     =   42246
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "ALL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   8655
      Left            =   -2880
      Picture         =   "frmSaleReport.frx":9CAA
      ScaleHeight     =   8595
      ScaleWidth      =   12915
      TabIndex        =   4
      Top             =   -1680
      Width           =   12975
      Begin VB.Shape Shape2 
         Height          =   1095
         Left            =   4560
         Shape           =   4  'Rounded Rectangle
         Top             =   5280
         Width           =   5415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "View Whole Sale Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   5400
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   5040
         TabIndex        =   8
         Top             =   2160
         Width           =   5535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Report According To Date"
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
         Left            =   4680
         TabIndex        =   7
         Top             =   3480
         Width           =   5415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   6
         Top             =   4080
         Width           =   375
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
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   4080
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   1695
         Left            =   4560
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmSaleReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAll_Click()
Call Connect
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strsql As String
strsql = "select * from Sales"
    With rs
   If rs.State = 1 Then .Close
    .Open strsql, con, adOpenDynamic, adLockBatchOptimistic
    If .RecordCount <> 0 Then
        
    Set SaleReport.DataSource = rs
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
strsql = "select * from Sales where Sale_Date BETWEEN #" & DTPickerFrom.Value & "# And #" & DTPickerTo.Value & "#"

    With rs
   If rs.State = 1 Then .Close
    .Open strsql, con, adOpenDynamic, adLockBatchOptimistic
    If .RecordCount <> 0 Then
            
        Set SaleReport.DataSource = rs
        SaleReport.Refresh
    Else
            MsgBox "No Record Available!", vbInformation, "Ambica"
           End If
           End With

End Sub

Private Sub Form_Load()
frmSaleReport.Top = 350
frmSaleReport.Left = 4100
'using 2 dtpickers
frmSaleReport.Left = 4100
DTPickerFrom.Value = Date
DTPickerTo.Value = Date

End Sub

