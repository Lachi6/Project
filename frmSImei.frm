VERSION 5.00
Begin VB.Form frmSImei 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmSImei.frx":0000
   ScaleHeight     =   5340
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSImei.frx":9CAA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtModelPrice 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtModelNo 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtImei 
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Left            =   0
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   600
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Searched Detail"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Model NO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "IMEI NO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblprice 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Left            =   4320
      TabIndex        =   7
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblname 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lbldate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblShow 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "frmSImei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim name1 As String

Private Sub Command1_Click()
Unload Me
End Sub

 Sub Form_Load()
name1 = InputBox("Please enter the IMEI to be searched.", vbOKOnly)
Call Connect
Set rs = New ADODB.Recordset
rs.Open "select * from IMEI where IMEI_NO like '" & name1 & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then
    
    Call Sale(name1)
Else
    Call showdata
End If
End Sub

Sub showdata()
txtImei.Text = rs.Fields(0)
txtModelNo.Text = rs.Fields(1)
txtModelPrice.Text = rs.Fields(2)
lblShow.Caption = "Unit In Stock"

End Sub
Sub showdatasale()
txtImei.Text = rs.Fields(2)
txtModelNo.Text = rs.Fields(3)
txtModelPrice.Text = rs.Fields(4)
lblShow.Caption = "Unit Sold on"
lbldate.Caption = rs.Fields(1).Value
lblName.Caption = rs.Fields(5).Value
lblprice.Caption = rs.Fields(7).Value
Label1.Caption = "To"
Label2.Caption = "At Rs"
End Sub

Sub Sale(name1)
If rs.State = 1 Then
rs.Close
End If

Set rs = New ADODB.Recordset
'rs.Open "select * from Sales where IMEI_NO='" & name1 & "'", con, adOpenDynamic, adLockOptimistic

 '  If rs.RecordCount > 0 Then
  '      Call showdatasale
  'Else
   '    lblShow.Caption = "No Record Found"
    '    Exit Sub
        
    'End If
rs.Open "select * from Sales", con, adOpenDynamic, adLockOptimistic

With rs
While Not .EOF
  If .Fields(2).Value = name1 Then
        Call showdatasale
    Else
        lblShow.Caption = "No record found"
     End If
.MoveNext
Wend

End With



End Sub

