Attribute VB_Name = "Module1"

Public Conn As New ADODB.Connection
Public RSKasir As ADODB.Recordset
Public RSAnggota As ADODB.Recordset
Public RSSimpan As ADODB.Recordset
Public RSPinjam As ADODB.Recordset

Public Sub BukaDB()
Set Conn = New ADODB.Connection
Set RSKasir = New ADODB.Recordset
Set RSAnggota = New ADODB.Recordset
Set RSSimpan = New ADODB.Recordset
Set RSPinjam = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBKoperasi.mdb"
End Sub



