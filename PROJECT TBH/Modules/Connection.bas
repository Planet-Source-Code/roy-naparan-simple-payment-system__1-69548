Attribute VB_Name = "Connection"
Public conn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public isLoginOk As Boolean

Option Explicit
Public Sub DBConnection()
   
        conn.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\Database\TBH.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=admin")
    
End Sub


