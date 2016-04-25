Attribute VB_Name = "Database"
Public con As New ADODB.Connection
Public rs As ADODB.Recordset
Public rs_Password As ADODB.Recordset
Public UserName As String
Public Rights As String
Public Status As String
Sub Connect()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs_Password = New ADODB.Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
         & App.Path & "\Tyre.mdb;Persist Security Info=False"
rs_Password.Open "Select * from Login ", con, adOpenDynamic, adLockOptimistic


End Sub


