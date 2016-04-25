VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000003&
   Caption         =   "Ambika Mobile Shop"
   ClientHeight    =   9885
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15240
   Icon            =   "MDImainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDImainForm.frx":2372
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      Picture         =   "MDImainForm.frx":1046B
      ScaleHeight     =   255
      ScaleWidth      =   15240
      TabIndex        =   2
      Top             =   840
      Width           =   15240
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome To Ambika Communication!!!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Top             =   0
         Width           =   16215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483639
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483628
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":1E564
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":200B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":21A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":23582
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":250D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":2B3ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":2CF3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":5CF91
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":5EAE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":60635
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImainForm.frx":90687
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1482
      ButtonWidth     =   2461
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Purchase Order"
            Key             =   "Porder"
            Object.ToolTipText     =   "Create Purchase Order"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sales "
            Key             =   "Sorder"
            Object.ToolTipText     =   "Sale Product"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stock In"
            Key             =   "Sin"
            Object.ToolTipText     =   "Stock In"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Suppliers"
            Key             =   "Supp"
            Object.ToolTipText     =   "Manage Suppliers"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Products"
            Key             =   "Prod"
            Object.ToolTipText     =   "Manage Products"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sales Details"
            Key             =   "SDetails"
            Object.ToolTipText     =   "View Sales Reports"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Purchase Details"
            Key             =   "PDetails"
            Object.ToolTipText     =   "View Purchase Report"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calculator"
            Key             =   "Calc"
            Object.ToolTipText     =   "Open Calculator"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Notepad"
            Key             =   "Notepad"
            Object.ToolTipText     =   "Open Notepad"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   7920
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   9525
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   988
            MinWidth        =   988
            Picture         =   "MDImainForm.frx":9284B
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Text            =   "User Name:"
            TextSave        =   "User Name:"
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
            Key             =   "User_name"
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1148
            MinWidth        =   1148
            Picture         =   "MDImainForm.frx":92C9D
            Object.ToolTipText     =   "Current Date and Time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "MDImainForm.frx":944D7
            Object.ToolTipText     =   "Current Date and Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Object.ToolTipText     =   "Current Date and Time"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "** Developed by :- Lachi Jain**"
            TextSave        =   "** Developed by :- Lachi Jain**"
            Object.ToolTipText     =   "Developer's Name"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Main"
      Begin VB.Menu mnuLogoff 
         Caption         =   "Log Off"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuChpass 
         Caption         =   "Change Password"
         Shortcut        =   ^C
      End
      Begin VB.Menu sdge 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Lock Application"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuCustSearch 
         Caption         =   "Search Customer"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnusearchImei 
         Caption         =   "Search By IMEI"
      End
      Begin VB.Menu mnuSearchSBN 
         Caption         =   "Supplier By Name"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuShortcut 
         Caption         =   "Shortcut Keys"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
      Begin VB.Menu AllRPT 
         Caption         =   "All Product"
      End
      Begin VB.Menu mnuPBS 
         Caption         =   "Product By Supplier"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administrator"
      Begin VB.Menu mnuUser 
         Caption         =   "User"
         Shortcut        =   ^S
      End
      Begin VB.Menu Model 
         Caption         =   "Model"
         Shortcut        =   {F6}
      End
      Begin VB.Menu dvg 
         Caption         =   "-"
      End
      Begin VB.Menu brand 
         Caption         =   "Brand"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Backup 
         Caption         =   "User Backup"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu feddevg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Us"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRpt As New ADODB.Recordset

Private Sub AllRPT_Click()
Call Connect

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim strsql As String
strsql = "select * from Model"
    With rs
   If rs.State = 1 Then .Close
        .Open strsql, con, adOpenDynamic, adLockBatchOptimistic
        If .RecordCount <> 0 Then
        
        Set DataReport1.DataSource = rs
        Else
            MsgBox "No Record Available!", vbInformation, "Ambica"
           End If
           End With

End Sub

Private Sub Backup_Click()
frmBackup.Show

End Sub

Private Sub brand_Click()
frmBrand.Show
End Sub

Private Sub mnuAbout_Click()
frmAbout1.Show
End Sub


Private Sub mnuCalculator_Click()
Shell "Calc.exe", vbNormalFocus

End Sub

Private Sub mnuChpass_Click()
frmChangePass.Show
End Sub

Private Sub mnuCustomer_Click()
frmCustomers.Show
End Sub

Private Sub mnuCustSearch_Click()
frmCustSearch.Show
End Sub

Private Sub mnuExit_Click()
If MsgBox("Are You Sure ?", vbYesNo + vbInformation, "Warning") = vbYes Then
    End
    Else
    Exit Sub
    End If
    End Sub

Private Sub mnuLock_Click()
frmLocked.Show vbModal

End Sub

Private Sub mnuLogOff_Click()
If MsgBox("Are you sure you want to Log Off?", vbYesNo + vbInformation, "Confirm Log Off") = vbYes Then
Unload Me
Unload MDIForm1
frmLogin.Show
Else
Exit Sub
End If
End Sub

Private Sub mnuNotepad_Click()
Shell "Notepad.exe", vbNormalFocus
End Sub

Private Sub mnuProducts_Click()
frmProduct.Show
End Sub

Private Sub mnuPurchaseD_Click()
Call showPurchasedetail
End Sub

Private Sub mnuPurchaseord_Click()
frmPurchaseOrd.Show
End Sub




Private Sub mnuSalesD_Click()
Call showSalesdetail
End Sub

Private Sub mnuSales_Click()
frmSale.Show
End Sub

Private Sub mnuPL_Click()
frmPL.Show
End Sub

Private Sub mnusearchImei_Click()
frmSImei.Show

End Sub

Private Sub mnuSearchSBN_Click()
frmSearchSBN.Show

End Sub

Private Sub mnuStockIn_Click()
frmStockIn.Show

End Sub



Private Sub mnuSupplier_Click()
frmSuppliers.Show

End Sub

Private Sub mnuShortcut_Click()
frmShortcut.Show
End Sub

Private Sub mnuUser_Click()
frmUser.Show

End Sub





Private Sub Model_Click()
frmModel.Show
End Sub

Private Sub mnuPBS_Click()
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

Private Sub Timer1_Timer()
StatusBar1.Panels(6).AutoSize = sbrContents
StatusBar1.Panels(6).Text = Format$(Date, "Long Date") & "  " & Format$(Time, "Long Time")

StatusBar1.Panels(8).Text = Right(StatusBar1.Panels(8).Text, Len(StatusBar1.Panels(8).Text) - 1) & Left(StatusBar1.Panels(8).Text, 1)

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key


Case "Porder"
frmPurchaseOrd.Show

Case "Sorder"
frmSale.Show

Case "Sin"
frmStockIn.Show
 
  
Case "Supp"
  frmSuppliers.Show
  
Case "Prod"
 frmProduct.Show

Case "PDetails"
  frmPurchaseReport.Show

Case "Calc"
Shell "Calc.exe", vbNormalFocus

Case "Notepad"
Shell "Notepad.exe", vbNormalFocus

Case "SDetails"
  frmSaleReport.Show
  

  
End Select
End Sub
