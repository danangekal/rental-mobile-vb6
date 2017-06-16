Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection
Public rslogin As New ADODB.Recordset
Public rsmerk As New ADODB.Recordset
Public rsmobil As New ADODB.Recordset
Public rsanggota As New ADODB.Recordset
Public rspenyewaan As New ADODB.Recordset
Public rspengembalian As New ADODB.Recordset

Public Sub koneksi()
Set conn = New ADODB.Connection
Set rslogin = New ADODB.Recordset
Set rsmerk = New ADODB.Recordset
Set rsmobil = New ADODB.Recordset
Set rsanggota = New ADODB.Recordset
Set rspenyewaan = New ADODB.Recordset
Set rspengembalian = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database1.mdb;Persist Security Info=False"
End Sub
