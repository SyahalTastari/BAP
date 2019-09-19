Imports System.Data.OleDb

Friend Class SqlCommand
    Private v As String
    Private cONN As OleDbConnection

    Public Sub New(v As String, cONN As OleDbConnection)
        Me.v = v
        Me.cONN = cONN
    End Sub
End Class
