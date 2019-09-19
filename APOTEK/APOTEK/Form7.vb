Public Class Form7
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggildata()
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_obat where nama_obat like '%" & TextBox8.Text & "%'", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_obat")
        DataGridView1.DataSource = DS.Tables("tb_obat")
    End Sub
    Sub panggildata()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_obat", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_obat")
        DataGridView1.DataSource = DS.Tables("tb_obat")
        DataGridView1.Enabled = True
    End Sub


    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Form2.tbkode_obat.Text = Me.DataGridView1.Item("kd_obat", i).Value
        Form2.tbnamaobat2.Text = Me.DataGridView1.Item("nama_obat", i).Value
        Form2.tbfungsi2.Text = Me.DataGridView1.Item("fungsi", i).Value
        Form2.tbstok2.Text = Me.DataGridView1.Item("stok", i).Value
        Form2.cmbjenis2.Text = Me.DataGridView1.Item("jenis", i).Value
        Form2.pbpasok.ImageLocation = Me.DataGridView1.Item("foto", i).Value
        Form2.pbpasok.SizeMode = PictureBoxSizeMode.StretchImage
        Me.Close()
    End Sub
End Class