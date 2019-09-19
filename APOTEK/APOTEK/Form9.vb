Public Class Form9
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggildata()
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_petugas where nama_petugas like '%" & TextBox8.Text & "%'", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_petugas")
        DataGridView1.DataSource = DS.Tables("tb_petugas")
    End Sub
    Sub panggildata()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_petugas", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_petugas")
        DataGridView1.DataSource = DS.Tables("tb_petugas")
        DataGridView1.Enabled = True
    End Sub


    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Form2.tbkode_petugas.Text = Me.DataGridView1.Item("kd_petugas", i).Value
        Form2.tbkode_petugas.Enabled = False
        Me.Close()
    End Sub
End Class