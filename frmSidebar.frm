VERSION 5.00
Begin VB.Form frmSidebar 
   BackColor       =   &H00808000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sidebar"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSidebar.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800080&
      FillColor       =   &H80000012&
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   3135
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      ToolTipText     =   "User's login time"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Time :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      ToolTipText     =   "User's Login time"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label lblUserName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "User name"
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "You Are Logged In As :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "User name"
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock In"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   14
      ToolTipText     =   "Enter Stock"
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick A Task"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   80
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   600
      Picture         =   "frmSidebar.frx":9CAA
      Stretch         =   -1  'True
      Top             =   75
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   12
      ToolTipText     =   "Exit From System"
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   11
      ToolTipText     =   "Change Your Current Password"
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Log Off"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      ToolTipText     =   "Log Off From System"
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Notepad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      ToolTipText     =   "Open Notepad"
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   8
      ToolTipText     =   "Open Calculator"
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "View Purchase Details"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      ToolTipText     =   "View Sales Details"
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Products"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "Manage Products"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Manage Suppliers"
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Customers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "Manage Customers"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "Take Purchase Order"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Give Purchase Order"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Out"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Remove Stock"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Image img13 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":CCBF
      ToolTipText     =   "Change your current password"
      Top             =   6480
      Width           =   480
   End
   Begin VB.Image img14 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":D989
      ToolTipText     =   "Exit from System"
      Top             =   6960
      Width           =   480
   End
   Begin VB.Image img2 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":E653
      ToolTipText     =   "Remove Stock"
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image img3 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":F31D
      ToolTipText     =   "Give Purchase Order"
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image img4 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":FFE7
      ToolTipText     =   "Take Sales Order"
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image img5 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":10CB1
      ToolTipText     =   "Manage Customers"
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image img6 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":1197B
      ToolTipText     =   "Manage Suppliers"
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image img7 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":12645
      ToolTipText     =   "Manage Products"
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image img8 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":1330F
      ToolTipText     =   "View Sales Details"
      Top             =   4080
      Width           =   480
   End
   Begin VB.Image img9 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":13FD9
      ToolTipText     =   "View Purchase Details"
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image img10 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":14CA3
      ToolTipText     =   "Open Calculator"
      Top             =   5040
      Width           =   480
   End
   Begin VB.Image img11 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":1596D
      ToolTipText     =   "Open Notepad"
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image img12 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":16637
      ToolTipText     =   "Log Off From System"
      Top             =   6000
      Width           =   480
   End
   Begin VB.Image img1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSidebar.frx":17301
      ToolTipText     =   "Enter Stock"
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "frmSidebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OrigColor As Long
Private OrigSize As Integer
Private OrigCaption As String

Private Sub Form_Load()
    OrigColor = Label2.ForeColor
    OrigSize = Label2.Font.Size
    OrigColor = Label3.ForeColor
    OrigSize = Label3.Font.Size
    OrigColor = Label4.ForeColor
    OrigSize = Label4.Font.Size
    OrigColor = Label5.ForeColor
    OrigSize = Label5.Font.Size
    OrigColor = Label6.ForeColor
    OrigSize = Label6.Font.Size
    OrigColor = Label7.ForeColor
    OrigSize = Label7.Font.Size
    OrigColor = Label8.ForeColor
    OrigSize = Label8.Font.Size
    OrigColor = Label9.ForeColor
    OrigSize = Label9.Font.Size
    OrigColor = Label10.ForeColor
    OrigSize = Label10.Font.Size
    OrigColor = Label11.ForeColor
    OrigSize = Label11.Font.Size
    OrigColor = Label12.ForeColor
    OrigSize = Label12.Font.Size
    OrigColor = Label13.ForeColor
    OrigSize = Label13.Font.Size
    OrigColor = Label14.ForeColor
    OrigSize = Label14.Font.Size
    OrigColor = Label15.ForeColor
    OrigSize = Label5.Font.Size
    OrigCaption = Label2.Caption
    lblTime.Caption = Time
    lblUserName.Caption = UserName
    
    
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Label2.Font.Size = OrigSize
       Label2.ForeColor = OrigColor
       Label3.Font.Size = OrigSize
       Label3.ForeColor = OrigColor
       Label4.Font.Size = OrigSize
       Label4.ForeColor = OrigColor
       Label5.Font.Size = OrigSize
       Label5.ForeColor = OrigColor
       Label6.Font.Size = OrigSize
       Label6.ForeColor = OrigColor
       Label7.Font.Size = OrigSize
       Label7.ForeColor = OrigColor
       Label8.Font.Size = OrigSize
       Label8.ForeColor = OrigColor
       Label9.Font.Size = OrigSize
       Label9.ForeColor = OrigColor
       Label10.Font.Size = OrigSize
       Label10.ForeColor = OrigColor
       Label11.Font.Size = OrigSize
       Label11.ForeColor = OrigColor
       Label12.Font.Size = OrigSize
       Label12.ForeColor = OrigColor
       Label13.Font.Size = OrigSize
       Label13.ForeColor = OrigColor
       Label14.Font.Size = OrigSize
       Label14.ForeColor = OrigColor
       Label15.Font.Size = OrigSize
       Label15.ForeColor = OrigColor
      Label2.Caption = OrigCaption
End Sub

Private Sub img1_Click()
frmStockIn.Show
End Sub

Private Sub img10_Click()
Shell "Calc.exe", vbNormalFocus
End Sub

Private Sub img11_Click()
Shell "Notepad.exe", vbNormalFocus
End Sub

Private Sub img12_Click()
If MsgBox("Are you sure you want to Log Off?", vbYesNo + vbInformation, "Confirm Log Off") = vbYes Then
Unload Me
Unload MDIForm1
frmLogin.Show
Else
Exit Sub
End If
End Sub

Private Sub img13_Click()
frmChangePass.Show
End Sub

Private Sub img14_Click()
If MsgBox("Are You Sure ?", vbYesNo + vbInformation, "Warning") = vbYes Then
    End
    Else
    Exit Sub
    End If
    
    
End
End Sub

Private Sub img2_Click()
frmStockOut.Show
End Sub

Private Sub img3_Click()
frmPurchaseOrd.Show
End Sub

Private Sub img4_Click()
frmSalesOrd.Show
End Sub

Private Sub img5_Click()
frmCustomers.Show
End Sub

Private Sub img6_Click()
frmSuppliers.Show
End Sub

Private Sub img7_Click()
frmProducts.Show
End Sub

Private Sub img8_Click()
Call showSalesdetail

End Sub

Private Sub img9_Click()
Call showPurchasedetail
End Sub

Private Sub Label10_Click()
img9_Click
End Sub

Private Sub Label11_Click()
img10_Click
End Sub

Private Sub Label12_Click()
img11_Click
End Sub

Private Sub Label13_Click()
img12_Click
End Sub

Private Sub Label14_Click()
img13_Click
End Sub

Private Sub Label15_Click()
img14_Click
End Sub

Private Sub Label2_Click()
img1_Click
End Sub

Private Sub Label3_Click()
img2_Click
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Label2.ForeColor = vbBlack
      Label2.Font.Size = 14
      'Label2.Caption = Right(Label2.Caption, Len(Label2.Caption) - 1) & Left(Label2.Caption, 1)
      
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbBlack
      Label3.Font.Size = 14
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlack
      Label4.Font.Size = 14
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label5.ForeColor = vbBlack
      Label5.Font.Size = 14
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlack
      Label6.Font.Size = 14
      
      
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbBlack
      Label7.Font.Size = 14
End Sub

Private Sub Label8_Click()
img7_Click
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = vbBlack
      Label8.Font.Size = 14
End Sub

Private Sub Label9_Click()
img8_Click
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = vbBlack
      Label9.Font.Size = 14
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = vbBlack
      Label10.Font.Size = 14
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = vbBlack
      Label11.Font.Size = 14
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = vbBlack
      Label12.Font.Size = 14
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = vbBlack
      Label13.Font.Size = 14
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = vbBlack
      Label14.Font.Size = 14
End Sub
Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.ForeColor = vbBlack
      Label15.Font.Size = 14
End Sub
Private Sub Label4_Click()
img3_Click
 
End Sub

Private Sub Label5_Click()
img4_Click
End Sub

Private Sub Label6_Click()
img5_Click
End Sub

Private Sub Label7_Click()
img6_Click
End Sub
