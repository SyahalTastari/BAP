Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim tbl As New DataTable
        konek()
        DA = New OleDb.OleDbDataAdapter("Select * From tb_petugas where username='" & TextBox1.Text & "'and password='" & TextBox2.Text & "'", CONN)
        DA.Fill(tbl)
        If tbl.Rows.Count > 0 Then
            MsgBox("Login berhasil.")
            Form2.lbnama.Text = tbl.Rows(0)(1).ToString + " sedang menggunakan"
            Form2.tbkode_petugas.Text = tbl.Rows(0)(0).ToString
            Form2.Show()
            Me.Hide()
        Else
            MsgBox("Anda tidak terdaftar.")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Ptterlihat_Click(sender As Object, e As EventArgs) Handles ptterlihat.Click
        TextBox2.UseSystemPasswordChar = True
        pttidakterlihat.Visible = True
        ptterlihat.Visible = False
    End Sub

    Private Sub pttidakterlihat_Click(sender As Object, e As EventArgs) Handles pttidakterlihat.Click
        TextBox2.UseSystemPasswordChar = False
        pttidakterlihat.Visible = False
        ptterlihat.Visible = True
    End Sub
End Class
