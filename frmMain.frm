VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000010&
   Caption         =   "Sales and Purchase Management System"
   ClientHeight    =   5820
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5460
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   988
            MinWidth        =   988
            Picture         =   "frmMain.frx":0000
            Object.ToolTipText     =   "User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Text            =   "User Name:"
            TextSave        =   "User Name:"
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1059
            MinWidth        =   1059
            Picture         =   "frmMain.frx":0452
            Object.ToolTipText     =   "Time of login"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Login Time:"
            TextSave        =   "Login Time:"
            Object.ToolTipText     =   "Time of login"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Time of login"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1148
            MinWidth        =   1148
            Picture         =   "frmMain.frx":08A4
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1325
            MinWidth        =   1325
            Text            =   "Date:"
            TextSave        =   "Date:"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "25/09/2015"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   988
            MinWidth        =   988
            Picture         =   "frmMain.frx":20DE
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:56 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   13215
      Left            =   0
      Picture         =   "frmMain.frx":2830
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuChPass 
         Caption         =   "Change password"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Add"
      Begin VB.Menu mnuAddC 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnuAddD 
         Caption         =   "Dealer"
      End
      Begin VB.Menu mnuAddP 
         Caption         =   "Product"
      End
      Begin VB.Menu mnuAddComp 
         Caption         =   "Company"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditC 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnuEditD 
         Caption         =   "Dealer"
      End
      Begin VB.Menu mnuEditP 
         Caption         =   "Product"
      End
      Begin VB.Menu mnuEditComp 
         Caption         =   "Company"
      End
   End
   Begin VB.Menu mnuSale 
      Caption         =   "&Sale"
   End
   Begin VB.Menu mnuPurchase 
      Caption         =   "&Purchase"
   End
   Begin VB.Menu mnuRecord 
      Caption         =   "&Record"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuSalesD 
         Caption         =   "Sales Details"
      End
      Begin VB.Menu mnuPurchaseD 
         Caption         =   "Purchase Details"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Begin VB.Menu mnuAddUser 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuRemoveUser 
         Caption         =   "Remove"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "A&bout Us"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
