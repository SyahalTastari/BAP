Imports System.Data.OleDb
Module Module1
    Public CONN As OleDbConnection
    Public CMD As OleDbCommand
    Public DA As OleDbDataAdapter
    Public RD As OleDbDataReader
    Public DS As New DataSet
    Public lokasidata As String
    Public Sub konek()
        lokasidata = "provider=microsoft.ace.oledb.12.0;data source= apotek.mdb"
        CONN = New OleDbConnection(lokasidata)
        If CONN.State = ConnectionState.Closed Then
            CONN.Open()
        End If
    End Sub
End Module
