Attribute VB_Name = "global"
Public obj As Control
Public prod_ID As String
Public i_Delete As String
Public iFind As String
Public ctr As Integer
Public i, j, k, l As Integer
Public poN As String


Public Prod_Code, Prod_Desc, Sp, Cat, U_Price, U_N_Stock As String
Public Sup_ID, Sup_Name, sAdd, sTel As String
Public scAd, itmAd, qtyAd As String

Public Sub CheckChar(KeyAscii As Integer)
    Dim a As Integer
    a = AllowChar(KeyAscii)
    KeyAscii = a
End Sub
Public Sub ValidNumeric(KeyAscii As Integer)
'allow only numeric value
'Check whether the Input is numeric or not
Select Case KeyAscii
Case 8
Case 48 To 57
Case 47
Case 13
Case 32
Case 48 To 57
 Case Else
  MsgBox "Invalid Input.Please Enter Numeric Types Only..", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select
End Sub
'allow only character value
Function AllowChar(a As Integer) As Integer 'doesn't allow you to enter special symbol
    Select Case a
        Case 65 To 90
        Case 97 To 122
        Case 8, 32
        Case Else
        a = 0
        MsgBox "Numeric field/Special Character is not allowed", vbInformation, "Invalid data"
    End Select
    AllowChar = a
End Function
Public Sub CheckspChar(KeyAscii As Integer)
    Dim a As Integer
    a = AllowChar(KeyAscii)
    KeyAscii = a
End Sub

' do not allow spesial character
Function DonotAllowSpChar(a As Integer) As Integer
Select Case a
    Case 65 To 90
    Case 97 To 122
    Case 8, 32
    Case 48 To 57
    Case Else
    a = 0
    MsgBox "Special Character is not allowed", vbInformation, "Invalid Input"
    End Select
    DonotAllowSpChar = a
 End Function
 
 Public Sub checkSpecialChar(KeyAscii As Integer)
   Dim a As Integer
   a = DonotAllowSpChar(KeyAscii)
   KeyAscii = a
 End Sub
 

