Public Class Form2
    Dim sqlnyaobat As String
    Dim sqlnyatransaksi As String
    Dim kot As String
    Dim jot As Integer
    Dim sot As Integer
    Dim bisa As Integer
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox8.Hide()
        pnlhome.BringToFront()
        btnhapuspetugas.Enabled = False
        btnupdateobat.Enabled = False
        tbpencarianobat.Visible = False
        tbpencarianpasok.Visible = False
        btnbatalpasok.Visible = False
        btnbatalpetugas.Visible = False
        tbpencarianpetugas.Visible = False
        btnhome.Visible = True
        btnhapusobat.Enabled = False
        btnbatalobat.Hide()
        tbpencariantransaksi.Visible = False
        btnupdatetransaksi.Enabled = True
        Call tampilkandataobat()
        Call panggildatapasok()
        Call panggildatapetugas()
        Call panggildatatransaksi()
        Call kodeautobat()
        Call kodeautopasok()
        Call kodeautopetugas()
        Call kodeautotransaksi()
        Dim btn As New DataGridViewButtonColumn()
        dgvtransaksi.Columns.Add(btn)
        btn.HeaderText = "PERSETUJUAN   "
        btn.Text = "BATAL"
        btn.Name = "btn"
        btn.UseColumnTextForButtonValue = True
    End Sub

    Sub kodeautotransaksi()
        konek()
        CMD = New OleDb.OleDbCommand("select * from tb_transaksi order by kd_transaksi desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()

        If Not RD.HasRows Then
            tbkodetransaksi.Text = "TI" + "001"
        Else
            tbkodetransaksi.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_transaksi").ToString, 4, 3)) + 1
            If Len(tbkodetransaksi.Text) = 1 Then
                tbkodetransaksi.Text = "TI00" & tbkodetransaksi.Text & ""
            ElseIf Len(tbkodetransaksi.Text) = 2 Then
                tbkodetransaksi.Text = "TI0" & tbkodetransaksi.Text & ""
            ElseIf Len(tbkodetransaksi.Text) = 3 Then
                tbkodetransaksi.Text = "TI" & tbkodetransaksi.Text & ""
            End If
        End If
    End Sub

    Sub kodeautobat()
        konek()
        CMD = New OleDb.OleDbCommand("select * from tb_obat order by kd_obat desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()

        If Not RD.HasRows Then
            tbkodeobat.Text = "OB" + "001"
        Else
            tbkodeobat.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_obat").ToString, 4, 3)) + 1
            If Len(tbkodeobat.Text) = 1 Then
                tbkodeobat.Text = "OB00" & tbkodeobat.Text & ""
            ElseIf Len(tbkodeobat.Text) = 2 Then
                tbkodeobat.Text = "OB0" & tbkodeobat.Text & ""
            ElseIf Len(tbkodeobat.Text) = 3 Then
                tbkodeobat.Text = "OB" & tbkodeobat.Text & ""
            End If
        End If
    End Sub
    Sub kembalikanstok()
        Dim sqlsss As String
        Dim stokss As Integer
        Dim i As Integer
        Dim a As Integer
        Dim b As Integer
        Dim c As String
        i = dgvtransaksi.CurrentRow.Index
        a = dgvtransaksi.Item("stok", i).Value
        b = dgvtransaksi.Item("jumlah_beli", i).Value
        c = dgvtransaksi.Item("kd_obat", i).Value
        stokss = a + b
        sqlsss = "update tb_obat set stok='" & stokss & "' where kd_obat='" & c & "'"
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlsss
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
    End Sub

    Sub kurangistok()
        Dim sqlss As String
        Dim stoks As Integer
        stoks = sot - jot
        sqlss = "update tb_obat set stok='" & stoks & "' where kd_obat='" & kot & "'"
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlss
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
    End Sub
    Sub kodeautopasok()
        konek()
        CMD = New OleDb.OleDbCommand("select * from tb_pasok order by kd_pasok desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            tbkodepasok.Text = "PK" + "001"
        Else
            tbkodepasok.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_pasok").ToString, 4, 3)) + 1
            If Len(tbkodepasok.Text) = 1 Then
                tbkodepasok.Text = "PK00" & tbkodepasok.Text & ""
            ElseIf Len(tbkodepasok.Text) = 2 Then
                tbkodepasok.Text = "PK0" & tbkodepasok.Text & ""
            ElseIf Len(tbkodepasok.Text) = 3 Then
                tbkodepasok.Text = "PK" & tbkodepasok.Text & ""
            End If
        End If
    End Sub

    Sub kodeautopetugas()
        konek()
        CMD = New OleDb.OleDbCommand("select * from tb_petugas order by kd_petugas desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            tbkodepetugas.Text = "PS" + "001"
        Else
            tbkodepetugas.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_petugas").ToString, 4, 3)) + 1
            If Len(tbkodepetugas.Text) = 1 Then
                tbkodepetugas.Text = "PS00" & tbkodepetugas.Text & ""
            ElseIf Len(tbkodepetugas.Text) = 2 Then
                tbkodepetugas.Text = "PS0" & tbkodepetugas.Text & ""
            ElseIf Len(tbkodepetugas.Text) = 3 Then
                tbkodepetugas.Text = "PS" & tbkodepetugas.Text & ""
            End If
        End If
    End Sub

    Sub panggildatatransaksi()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT kd_transaksi,kd_obat,kd_petugas,nama_obat,harga,jumlah_beli,subtotal,stok,tgl_transaksi FROM qw_transaksi where kd_transaksi='" & tbkodetransaksi.Text & "'", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "qw_transaksi")
        dgvtransaksi.DataSource = DS.Tables("qw_transaksi")
        dgvtransaksi.Enabled = True
    End Sub

    Sub jalantransaksi()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnyatransaksi
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        'tbkodetransaksi.Text = ""
        dtptransaksi.ResetText()
        tbjumlahbeli.Text = ""
        'tbsubtotal.Text = ""
        tbnama__obat.Text = ""
        tbkode__obat.Text = ""
        tbharga.Text = ""
        'tbkode_petugas.Text = ""
        tbstoktransaksi.Text = ""
    End Sub

    Sub panggildatapetugas()
        Call konek()
        DA = New OleDb.OleDbDataAdapter("SELECT kd_petugas,nama_petugas,jk,no_hp,alamat,username from tb_petugas", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_petugas")
        dgvpetugas.DataSource = DS.Tables("tb_petugas")
        dgvpetugas.ReadOnly = True
    End Sub
    Sub stok()
        Dim sqls As String
        Dim stock As Integer
        stock = Val(tbjumlahpasok.Text) + Val(tbstok2.Text)
        sqls = "update tb_obat set stok='" & stock & "' where kd_obat='" & tbkode_obat.Text & "'"
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqls
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
    End Sub
    Dim sqlnyastruk As String
    Sub jalanstruk()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnyastruk
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
    End Sub
    Sub jalanobat()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnyaobat
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        tbkodeobat.Text = ""
        tbnamaobat.Text = ""
        cmbjenis.Text = ""
        cmbjenis.SelectedIndex = -1
        cmbkemasan.Text = ""
        cmbkemasan.SelectedIndex = -1
        tbfungsi.Text = ""
        tbstok.Text = ""
        tbhargaobat.Text = ""
        tbpencarianobat.Text = ""
        tblokasigambar.Text = ""
        fbfoto.ImageLocation = Nothing
    End Sub
    Dim sqlnyapetugas As String
    Sub jalanpetugas()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnyapetugas
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        tbkodepetugas.Text = ""
        tbnama.Text = ""
        tbtlp.Text = ""
        tbalamat.Text = ""
        tbusername.Text = ""
        tbpassword.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub btnkelolaobat_Click(sender As Object, e As EventArgs) Handles btnkelolaobat.Click
        pnlkelolaobat.BringToFront()
        Call tampilkandataobat()
        btnhome.Visible = True
    End Sub
    Private Sub Btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        pnlhome.BringToFront()
    End Sub
    Private Sub Btnkelolapasok_Click(sender As Object, e As EventArgs)
        pnlpasokobat.BringToFront()
        Call panggildatapasok()
        btnhome.Visible = True
    End Sub
    Private Sub Label41_Click(sender As Object, e As EventArgs)
        pnlhome.BringToFront()
    End Sub

    Private Sub Btnkelolapetugas_Click(sender As Object, e As EventArgs) Handles btnkelolapetugas.Click
        pnlkelolapetugas.BringToFront()
        btnhome.Visible = True
    End Sub

    Private Sub Btntransaksi_Click(sender As Object, e As EventArgs) Handles btntransaksi.Click
        pnltransaksi.BringToFront()
        btnhome.Visible = True
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)
        pnlhome.BringToFront()
    End Sub
    Private Sub Btncarikodeobat_Click(sender As Object, e As EventArgs) Handles btncarikodeobat.Click
        Dim a As String
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = "select * from tb_obat where kd_obat='" & tbkode__obat.Text & "'"
        RD = objcmd.ExecuteReader()
        RD.Read()
        a = MsgBox("Apakah Anda Ingin Melihat Data Barang?", vbYesNo, "Medica")
        If a = vbYes Then
            Form8.Show()
            tbjumlahbeli.Text = ""
        End If
    End Sub
    Sub tampilkandataobat()
        Call konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * from tb_obat", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_obat")
        dgvkelolaobat.DataSource = DS.Tables(0)
        dgvkelolaobat.ReadOnly = True
    End Sub

    Private Sub Btnsimpan_Click(sender As Object, e As EventArgs) Handles btnsimpanobat.Click
        If tbkodeobat.Text = "" Or tbnamaobat.Text = "" Or cmbjenis.Text = "" Or cmbkemasan.Text = "" Or tbnamaobat.Text = "" Or tbhargaobat.Text = "" Then
            MsgBox("Tolong Isi Semua Data")
        Else
            sqlnyaobat = "insert into tb_obat (kd_obat,nama_obat,jenis,kemasan,fungsi,stok,foto,harga) values('" & tbkodeobat.Text & "','" & tbnamaobat.Text & "','" & cmbjenis.Text & "','" & cmbkemasan.Text & "','" & tbfungsi.Text & "','" & tbstok.Text & "' , '" & tblokasigambar.Text & "' , '" & tbhargaobat.Text & "')"
            Call jalanobat()
            MsgBox("Data Berhasil Tersimpan")
            Call tampilkandataobat()
            Call kodeautobat()
            fbfoto.Image = Nothing
            tbstok.Text = 0
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub Btnhapus_Click(sender As Object, e As EventArgs) Handles btnhapusobat.Click
        sqlnyaobat = "delete from tb_obat where kd_obat='" & tbkodeobat.Text & "'"
        Call jalanobat()
        tbstok.Text = 0
        MsgBox("Data Berhasil Terhapus")
        Call tampilkandataobat()
        Call kodeautobat()
        tbkodeobat.Enabled = True
    End Sub

    Private Sub dgvkelolaobat_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvkelolaobat.RowHeaderMouseClick
        Dim i As Integer
        i = dgvkelolaobat.CurrentRow.Index
        If i = dgvkelolaobat.NewRowIndex Then
            MsgBox("Tidak Bisa Memilih Data Kosong")
        Else
            tbkodeobat.Text = dgvkelolaobat.Item(0, i).Value
            tbnamaobat.Text = dgvkelolaobat.Item(1, i).Value
            cmbjenis.Text = dgvkelolaobat.Item(2, i).Value
            cmbkemasan.Text = dgvkelolaobat.Item(3, i).Value
            tbfungsi.Text = dgvkelolaobat.Item(4, i).Value
            tbstok.Text = dgvkelolaobat.Item(5, i).Value
            tblokasigambar.Text = dgvkelolaobat.Item(6, i).Value
            fbfoto.ImageLocation = dgvkelolaobat.Item(6, i).Value
            fbfoto.SizeMode = PictureBoxSizeMode.StretchImage
            tbhargaobat.Text = dgvkelolaobat.Item(7, i).Value
            tbkodeobat.Enabled = False
            btnsimpanobat.Enabled = False
            btnupdateobat.Enabled = True
            btnhapusobat.Enabled = True
            btnbatalobat.Show()
        End If
    End Sub

    Private Sub Btnupdate_Click(sender As Object, e As EventArgs) Handles btnupdateobat.Click
        If tbkodeobat.Text = "" Or tbnamaobat.Text = "" Or cmbjenis.Text = "" Or cmbkemasan.Text = "" Or tbfungsi.Text = "" Or tbstok.Text = "" Then
            MsgBox("Tolong Lengkapi Data")
        Else
            sqlnyaobat = "UPDATE tb_obat set nama_obat='" & tbnamaobat.Text & "',jenis='" & cmbjenis.Text & "',kemasan='" & cmbkemasan.Text & "',fungsi='" & tbfungsi.Text & "',stok='" & tbstok.Text & "',foto='" & tblokasigambar.Text & "',harga='" & tbhargaobat.Text & "'where kd_obat='" & tbkodeobat.Text & "'"
            Call jalanobat()
            MsgBox("Data Berhasil Di Update")
            Call tampilkandataobat()
            Call kodeautobat()
            btnbatalobat.Visible = False
            fbfoto.Image = Nothing
            tbkodeobat.Enabled = True
            btnsimpanobat.Enabled = True
            tbstok.Text = 0

        End If

    End Sub

    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Me.Close()
        Form1.Close()
    End Sub

    Private PathFile As String = Nothing
    Private Sub Btncari_Click(sender As Object, e As EventArgs) Handles btncari.Click
        Dim gambar As String
        OpenFileDialog2.Filter = "JPG Files(*.jpg)|*.jpg|JPEG Files (*.jpeg)|*.jpeg|GIF Files(*.gif)|*.gif|PNG Files(*.png)|*.png|BMP Files(*.bmp)|*.bmp|TIFF Files(*.tiff)|*.tiff"
        OpenFileDialog2.FileName = ""
        If OpenFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
            fbfoto.SizeMode = PictureBoxSizeMode.StretchImage
            fbfoto.Image = New Bitmap(OpenFileDialog2.FileName)
            btncari.Enabled = True
            PathFile = OpenFileDialog2.FileName
            TextBox8.Text = PathFile.Substring(PathFile.LastIndexOf("\") + 1)
            tblokasigambar.Text = OpenFileDialog2.FileName
            gambar = OpenFileDialog2.FileName
            fbfoto.Image = Image.FromFile(tblokasigambar.Text)
        End If
    End Sub

    Private Sub Btnbatalobat_Click_1(sender As Object, e As EventArgs) Handles btnbatalobat.Click
        tbkodeobat.Enabled = True
        Dim kon As String
        kon = MsgBox("Batalkan Untuk Memilih?", vbYesNo, "Sistem Lab Say")
        If kon = vbYes Then
            tbkodeobat.Text = ""
            tbnamaobat.Text = ""
            cmbjenis.Text = ""
            cmbjenis.SelectedIndex = -1
            cmbkemasan.Text = ""
            cmbkemasan.SelectedIndex = -1
            tbfungsi.Text = ""
            tbstok.Text = ""
            tbhargaobat.Text = ""
            fbfoto.Image = Nothing
            tblokasigambar.Text = ""
            btnhapusobat.Enabled = False
            btnupdateobat.Enabled = False
            btnsimpanobat.Enabled = True
            tbkodeobat.Enabled = True
            btnbatalobat.Hide()
            Call kodeautobat()
            tbstok.Text = 0
        Else
        End If
    End Sub

    Private Sub Tbpencarian_TextChanged(sender As Object, e As EventArgs) Handles tbpencarianobat.TextChanged
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_obat where nama_obat like '%" & tbpencarianobat.Text & "%'", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_obat")
        dgvkelolaobat.DataSource = DS.Tables("tb_obat")
    End Sub

    Private Sub BunifuImageButton2_Click(sender As Object, e As EventArgs) Handles BunifuImageButton2.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub BunifuImageButton3_Click(sender As Object, e As EventArgs) Handles BunifuImageButton3.Click
        tbpencarianobat.Visible = True
    End Sub

    Private Sub btnupdatepasok_Click(sender As Object, e As EventArgs) Handles btnupdatepasok.Click
        If tbkodepasok.Text = "" Or tbkode_obat.Text = "" Or tbjumlahpasok.Text = "" Or tb_distributor.Text = "" Then
            MsgBox("Lengkapi Data")
        Else
            sqlnyapasok = "UPDATE qw_pasok set kd_obat='" & tbkode_obat.Text & "',nama_obat='" & tbnamaobat2.Text & "',jenis='" & cmbjenis2.Text & "',fungsi='" & tbfungsi2.Text & "',jumlah_pasok='" & tbjumlahpasok.Text & "',distributor='" & tb_distributor.Text & "',tanggal_pasok='" & dtpasok.Text & "' where kd_pasok='" & tbkodepasok.Text & "'"
            Call jalanpasok()
            MsgBox("Data Berhasil Terubah")
            Call panggildatapasok()
            btnsimpanpasok.Enabled = True
            btnhapuspasok.Enabled = False
            btnupdatepasok.Enabled = False
            btnbatalpasok.Visible = False
            tbkodepasok.Enabled = True
            Call kodeautopasok()
        End If
    End Sub

    Private Sub BunifuImageButton4_Click(sender As Object, e As EventArgs) Handles btncaripasok.Click
        tbpencarianpasok.Visible = True
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles btnsimpanpasok.Click
        If tbkodepasok.Text = "" Or tbkode_obat.Text = "" Or tbjumlahpasok.Text = "" Or tb_distributor.Text = "" Or tbnamaobat2.Text = "" Or tbfungsi2.Text = "" Or cmbjenis2.Text = "" Then
            MsgBox("Tolong Isi Semua Data")
        Else
            sqlnyapasok = "insert into tb_pasok(`kd_pasok`,`kd_obat`,`jumlah_pasok`,`tanggal_pasok`,`distributor`) values('" & tbkodepasok.Text & "','" & tbkode_obat.Text & "','" & tbjumlahpasok.Text & "','" & dtpasok.Text & "','" & tb_distributor.Text & "')"
            Call stok()
            Call jalanpasok()
            MsgBox("Data Berhasil Tersimpan")
            Call panggildatapasok()
            Call kodeautopasok()
        End If
        btnsimpanpasok.Enabled = True
        btnhapuspasok.Enabled = False
        btnupdatepasok.Enabled = False
        btnbatalpasok.Visible = False
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form7.Show()
    End Sub


    Sub panggildatapasok()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM qw_pasok", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "qw_pasok")
        dgvpasok.DataSource = DS.Tables("qw_pasok")
        dgvpasok.Enabled = True
    End Sub
    Dim sqlnyapasok As String
    Sub jalanpasok()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnyapasok
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        tbkodepasok.Text = ""
        tbkode_obat.Text = ""
        tbjumlahpasok.Text = ""
        dtpasok.ResetText()
        tb_distributor.Text = ""
        tbnamaobat2.Text = ""
        tbfungsi2.Text = ""
        tbstok2.Text = ""
        cmbjenis2.Text = ""
        cmbjenis2.SelectedIndex = -1
        pbpasok.Image = Nothing
    End Sub


    Private Sub dgvpasok_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvpasok.RowHeaderMouseClick
        Dim i As Integer
        i = dgvpasok.CurrentRow.Index
        If i = dgvkelolaobat.NewRowIndex Then
            MsgBox("Tidak Bisa Memilih Data Kosong")
        Else
            btnbatalobat.Visible = True
            tbkodepasok.Text = dgvpasok.Item("kd_pasok", i).Value
            tbkode_obat.Text = dgvpasok.Item("kd_obat", i).Value
            tbjumlahpasok.Text = dgvpasok.Item("jumlah_pasok", i).Value
            tb_distributor.Text = dgvpasok.Item("distributor", i).Value
            dtpasok.Text = dgvpasok.Item("tanggal_pasok", i).Value
            Dim objcmd As New System.Data.OleDb.OleDbCommand
            Call konek()
            objcmd.Connection = CONN
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = "select * from tb_obat where kd_obat='" & tbkode_obat.Text & "'"
            RD = objcmd.ExecuteReader()
            RD.Read()
            If RD.HasRows Then
                tbnamaobat2.Text = RD.Item("nama_obat")
                tbstok2.Text = RD.Item("stok")
                cmbjenis2.Text = RD.Item("jenis")
                tbfungsi2.Text = RD.Item("fungsi")
                pbpasok.ImageLocation = RD.Item("foto")
                pbpasok.SizeMode = PictureBoxSizeMode.StretchImage
            End If
            btnsimpanpasok.Enabled = False
            btnhapuspasok.Enabled = True
            btnupdatepasok.Enabled = True
            btnbatalpasok.Visible = True
            tbkodepasok.Enabled = False
        End If
    End Sub

    Private Sub Btnrefresh_Click(sender As Object, e As EventArgs)
        Call tampilkandataobat()
    End Sub

    Private Sub Btnhapuspasok_Click(sender As Object, e As EventArgs) Handles btnhapuspasok.Click
        sqlnyapasok = "delete from tb_pasok where kd_pasok='" & tbkodepasok.Text & "'"
        Call jalanpasok()
        MsgBox("Data Berhasil Terhapus")
        Call panggildatapasok()
        btnsimpanpasok.Enabled = True
        btnhapuspasok.Enabled = False
        btnupdatepasok.Enabled = False
        btnbatalpasok.Visible = False
        tbkodepasok.Enabled = True
        Call kodeautopasok()
    End Sub

    Private Sub Tbpencarianpasok_TextChanged(sender As Object, e As EventArgs) Handles tbpencarianpasok.TextChanged
        konek()
        DA = New OleDb.OleDbDataAdapter("select * from qw_pasok where nama_obat like '%" & tbpencarianpasok.Text & "%'", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "qw_pasok")
        dgvpasok.DataSource = DS.Tables("qw_pasok")
    End Sub

    Private Sub Btnbatalpasok_Click(sender As Object, e As EventArgs) Handles btnbatalpasok.Click
        tbkodepasok.Enabled = True
        Dim kon As String
        kon = MsgBox("Tidak Jadi Mengubah atau Menghapus?", vbYesNo, "Medika Care")
        If kon = vbYes Then
            tbkodepasok.Text = ""
            tbkode_obat.Text = ""
            tbjumlahpasok.Text = ""
            tb_distributor.Text = ""
            tbnamaobat2.Text = ""
            tbfungsi2.Text = ""
            tbstok2.Text = ""
            dtpasok.Value = Date.Now
            cmbjenis2.Text = ""
            cmbjenis2.SelectedIndex = -1
            pbpasok.Image = Nothing
            btnsimpanpasok.Enabled = True
            btnhapuspasok.Enabled = False
            btnupdatepasok.Enabled = False
            btnbatalpasok.Visible = False
            tbkodepasok.Enabled = True
            Call kodeautopasok()
        Else
        End If
    End Sub

    '===================================================PANEL PETUGAS=============================================================='

    Private Sub BunifuImageButton4_Click_1(sender As Object, e As EventArgs) Handles BunifuImageButton4.Click
        tbpencarianpetugas.Visible = True

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles btnsimpanpetugas.Click
        Dim jk As String
        jk = ""
        If RadioButton1.Checked = True Then
            jk = "Laki-Laki"
        ElseIf RadioButton2.Checked = True Then
            jk = "Perempuan"
        End If
        If tbkodepetugas.Text = "" Or tbnama.Text = "" Or tbtlp.Text = "" Or tbtlp.Text = "" Or tbalamat.Text = "" Or tbusername.Text = "" Or tbpassword.Text = "" Or (RadioButton1.Checked = False And RadioButton2.Checked = False) Then
            MsgBox("Tolong Isi Semua Data")
        Else
            Call konek()
            CMD = New OleDb.OleDbCommand("Select * from tb_petugas where username='" & tbusername.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                sqlnyapetugas = "insert into tb_petugas(`kd_petugas`,`nama_petugas`,`jk`,`no_hp`,`alamat`,`username`,`password`) values('" & tbkodepetugas.Text & "','" & tbnama.Text & "','" & jk & "','" & tbtlp.Text & "','" & tbalamat.Text & "','" & tbusername.Text & "','" & tbpassword.Text & "')"
                Call jalanpetugas()
                MsgBox("Data Berhasil Tersimpan")
                Call panggildatapetugas()
            Else
                MsgBox("username sudah ada")
            End If
            btnhapuspetugas.Enabled = False
            btnupdatepetugas.Enabled = False
            Call kodeautopetugas()
        End If
    End Sub
    Private Sub Btnhapuspetugas_Click(sender As Object, e As EventArgs) Handles btnhapuspetugas.Click
        sqlnyapetugas = "delete from tb_petugas where kd_petugas='" & tbkodepetugas.Text & "'"
        Call jalanpetugas()
        MsgBox("Data Berhasil Terhapus")
        Call panggildatapetugas()
        btnsimpanpetugas.Enabled = True
        btnhapuspetugas.Enabled = False
        btnupdatepetugas.Enabled = False
        btnbatalpetugas.Visible = False
        tbkodepetugas.Enabled = True
        Call kodeautopetugas()
    End Sub

    Private Sub dgvpetugas_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvpetugas.RowHeaderMouseClick
        Dim i As Integer
        i = dgvpetugas.CurrentRow.Index
        If i = dgvkelolaobat.NewRowIndex Then
            MsgBox("Tidak Bisa Memilih Data Kosong")
        Else

            tbkodepetugas.Text = dgvpetugas.Item(0, i).Value
            tbnama.Text = dgvpetugas.Item(1, i).Value
            If dgvpetugas.Item(2, i).Value = "Laki-Laki" Then
                RadioButton1.Checked = True
            ElseIf dgvpetugas.Item(2, i).Value = "Perempuan" Then
                RadioButton2.Checked = True
            End If
            tbtlp.Text = dgvpetugas.Item(3, i).Value
            tbalamat.Text = dgvpetugas.Item(4, i).Value
            tbusername.Text = dgvpetugas.Item(5, i).Value
            tbpassword.Text = dgvpetugas.Item(6, i).Value
        End If
        btnsimpanpetugas.Enabled = False
        btnhapuspetugas.Enabled = True
        btnupdatepetugas.Enabled = True
        btnbatalpetugas.Visible = True
        tbkodepetugas.Enabled = False
    End Sub

    Private Sub Btnupdatepetugas_Click(sender As Object, e As EventArgs) Handles btnupdatepetugas.Click
        Dim jks As String
        jks = ""
        If RadioButton1.Checked = True Then
            jks = "Laki-Laki"
        ElseIf RadioButton2.Checked = True Then
            jks = "Perempuan"
        End If
        If tbkodepetugas.Text = "" Or tbnama.Text = "" Or tbtlp.Text = "" Or (RadioButton1.Checked = False And RadioButton2.Checked = False) Then
            MsgBox("Tolong Isi Semua Data")
        Else
            sqlnyapetugas = "UPDATE tb_petugas set [nama_petugas] = '" & tbnama.Text & "',[jk] = '" & jks & "',[no_hp] = '" & tbtlp.Text & "',[alamat] = '" & tbalamat.Text & "',[username] = '" & tbusername.Text & "',[password] = '" & tbpassword.Text & "' where [kd_petugas] = '" & tbkodepetugas.Text & "'"
            Call jalanpetugas()
            MsgBox("Data Berhasil Terubah")
            Call panggildatapetugas()
            btnsimpanpetugas.Enabled = True
            btnhapuspetugas.Enabled = False
            btnupdatepetugas.Enabled = False
            btnbatalpetugas.Visible = False
            Call kodeautopetugas()
        End If
    End Sub

    Private Sub Btnbatalpetugas_Click(sender As Object, e As EventArgs) Handles btnbatalpetugas.Click
        tbkodepetugas.Enabled = True
        Dim ch As String
        ch = MsgBox("Tidak Jadi Mengubah atau Menghapus?", vbYesNo, "Medika Care")
        If ch = vbYes Then
            tbkodepetugas.Text = ""
            tbnama.Text = ""
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            tbtlp.Text = ""
            tbalamat.Text = ""
            tbusername.Text = ""
            tbpassword.Text = ""
            btnsimpanpetugas.Enabled = True
            btnhapuspetugas.Enabled = False
            btnupdatepetugas.Enabled = False
            btnbatalpetugas.Visible = False
        Else
        End If
    End Sub

    Private Sub Tbpencarianpetugas_TextChanged(sender As Object, e As EventArgs) Handles tbpencarianpetugas.TextChanged
        konek()
        DA = New OleDb.OleDbDataAdapter("select * from tb_petugas where nama_petugas like '%" & tbpencarianpetugas.Text & "%'", CONN)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_petugas")
        dgvpetugas.DataSource = DS.Tables("tb_petugas")
    End Sub
    '========================================TRANSKAKSI============================================================='
    Private Sub Btncarikodepetugas_Click(sender As Object, e As EventArgs)
        Dim a As String
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = CONN
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = "select * from tb_obat where kd_obat='" & tbkode__obat.Text & "'"
        RD = objcmd.ExecuteReader()
        RD.Read()
        a = MsgBox("Apakah Anda Ingin Melihat Data Petugas?", vbYesNo, "Medica")
        If a = vbYes Then
            Form9.Show()
        End If
    End Sub

    Private Sub Tbjumlahbeli_TextChanged(sender As Object, e As EventArgs) Handles tbjumlahbeli.TextChanged
        tbsubtotal.Text = Val(tbjumlahbeli.Text) * Val(tbharga.Text)
    End Sub
    Private Sub Btnsimpantransaksi_Click(sender As Object, e As EventArgs) Handles btnsimpantransaksi.Click
        If tbkodetransaksi.Text = "" Or tbharga.Text = "" Or tbsubtotal.Text = "" Or tbkode__obat.Text = "" Or tbnama__obat.Text = "" Or tbharga.Text = "" Or tbkode_petugas.Text = "" Then
            MsgBox("Tolong Isi Semua Data")
        ElseIf tbstoktransaksi.Text = 0 Then
            MsgBox("Maaf stok sedang tidak tersedia ")
        Else
            Call konek()
            CMD = New OleDb.OleDbCommand("Select * from tb_transaksi where kd_transaksi='" & tbkodetransaksi.Text & "' and kd_obat='" & tbkode__obat.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                kot = tbkode__obat.Text
                jot = tbjumlahbeli.Text
                sot = tbstoktransaksi.Text
                sqlnyatransaksi = "insert into tb_transaksi(`kd_transaksi`,`kd_obat`,`kd_petugas`,`jumlah_beli`,`subtotal`,`tgl_transaksi`) values('" & tbkodetransaksi.Text & "','" & kot & "','" & tbkode_petugas.Text & "','" & jot & "','" & tbsubtotal.Text & "','" & dtptransaksi.Text & "')"
                Call kurangistok()
                Call jalantransaksi()
                tbjumlahbeli.Enabled = False
                MsgBox("Data Berhasil Tersimpan")
                Call panggildatatransaksi()
            Else
                MsgBox("Maaf kode sudah ada, Jika anda ingin memasukan kode yang sama anda harus membatalkan data yang tercantum di tabel")
            End If
        End If
        tbsubtotal.Text = 0
        Button1.Enabled = True


    End Sub

    Private Sub tbtlp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbtlp.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled = True
    End Sub

    Private Sub dgvtransaksi_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvtransaksi.CellContentClick
        Dim x, y, m As String
        Dim i As Integer
        i = dgvtransaksi.CurrentRow.Index
        If e.ColumnIndex = 0 Then
            x = dgvtransaksi.Item("kd_obat", i).Value
            y = dgvtransaksi.Item("kd_transaksi", i).Value
            m = MsgBox("Yakin Untuk Membatalkan ?", vbYesNo, "Medica")
            If m = vbYes Then
                sqlnyatransaksi = "delete from tb_transaksi where kd_transaksi='" & y & "' and kd_obat='" & x & "'"
                Call kembalikanstok()
                Call jalantransaksi()
                MsgBox("data berhasil dihapus")
                Call panggildatatransaksi()
            End If
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sum As Integer = 0
        If dgvtransaksi.CurrentCell Is Nothing Then
            MsgBox("Tidak dapat mentotalkan jika tabel masih kosong")
        Else
            For i As Integer = 0 To dgvtransaksi.Rows.Count() - 1 Step +1
                sum = sum + dgvtransaksi.Rows(i).Cells(7).Value
            Next
            tbtotal.Text = sum.ToString
        End If
    End Sub
    Sub pemecahauangkembalian()
        tbkembalian.Text = Val(tbbayar.Text) - Val(tbtotal.Text)
        Dim jmluang As Integer
        Dim sisauang As Integer
        Dim jml100rb As Double
        Dim sisa100rb As Integer
        Dim jml50rb As Double
        Dim sisa50rb As Integer
        Dim jml20rb As Double
        Dim sisa20rb As Integer
        Dim jml10rb As Double
        Dim sisa10rb As Integer
        Dim jml5rb As Double
        Dim sisa5rb As Integer
        Dim jml2rb As Double
        Dim sisa2rb As Integer
        Dim jml1rb As Double
        Dim sisa1rb As Integer
        Dim jml500 As Double
        Dim sisa500 As Integer
        jmluang = tbkembalian.Text
        sisauang = jmluang Mod 25
        If sisauang > 0 Then
            MsgBox("Jumlah uang salah")
        Else
            jml100rb = jmluang / 100000
            txt100rb.Text = Math.Floor(jml100rb)
            sisa100rb = jmluang - (100000 * Int(jml100rb))
            jml50rb = sisa100rb / 50000
            txt50rb.Text = Math.Floor(jml50rb)
            sisa50rb = sisa100rb - (50000 * Int(jml50rb))
            jml20rb = sisa50rb / 20000
            txt20rb.Text = Math.Floor(jml20rb)
            sisa20rb = sisa50rb - (20000 * Int(jml20rb))
            jml10rb = sisa20rb / 10000
            txt10rb.Text = Math.Floor(jml10rb)
            sisa10rb = sisa20rb - (10000 * Int(jml10rb))
            jml5rb = sisa10rb / 5000
            txt5rb.Text = Math.Floor(jml5rb)
            sisa5rb = sisa10rb - (5000 * Int(jml5rb))
            jml2rb = sisa5rb / 2000
            txt2rb.Text = Math.Floor(jml2rb)
            sisa2rb = sisa5rb - (2000 * Int(jml2rb))
            jml1rb = sisa2rb / 1000
            txt1rb.Text = Math.Floor(jml1rb)
            sisa1rb = sisa2rb - (1000 * Int(jml1rb))
            jml500 = sisa1rb / 500
            txt5rts.Text = Math.Floor(jml500)
            sisa500 = sisa1rb - (500 * Int(jml500))
        End If
    End Sub
    Dim a As Integer
    Private Sub Btnupdatetransaksi_Click(sender As Object, e As EventArgs) Handles btnupdatetransaksi.Click
        If Val(tbbayar.Text) >= Val(tbtotal.Text) Then
            Call pemecahauangkembalian()
        Else
            MsgBox("Uang Anda Tidak Cukup")
        End If
    End Sub
    Sub pecahanuang()
        Dim jmluang As Integer
        Dim sisauang As Integer
        Dim jml100rb As Double
        Dim sisa100rb As Integer
        Dim jml50rb As Double
        Dim sisa50rb As Integer
        Dim jml20rb As Double
        Dim sisa20rb As Integer
        Dim jml10rb As Double
        Dim sisa10rb As Integer
        Dim jml5rb As Double
        Dim sisa5rb As Integer
        Dim jml2rb As Double
        Dim sisa2rb As Integer
        Dim jml1rb As Double
        Dim sisa1rb As Integer
        Dim jml500 As Double
        Dim sisa500 As Integer
        jmluang = Val(tbtotal.Text)
        sisauang = jmluang Mod 25
        If sisauang > 0 Then
            MsgBox("Jumlah uang salah")
        Else
            jml100rb = jmluang / 100000
            txt100rb.Text = Math.Floor(jml100rb)
            sisa100rb = jmluang - (100000 * Int(jml100rb))
            jml50rb = sisa100rb / 50000
            txt50rb.Text = Math.Floor(jml50rb)
            sisa50rb = sisa100rb - (50000 * Int(jml50rb))
            jml20rb = sisa50rb / 20000
            txt20rb.Text = Math.Floor(jml20rb)
            sisa20rb = sisa50rb - (20000 * Int(jml20rb))
            jml10rb = sisa20rb / 10000
            txt10rb.Text = Math.Floor(jml10rb)
            sisa10rb = sisa20rb - (10000 * Int(jml10rb))
            jml5rb = sisa10rb / 5000
            txt5rb.Text = Math.Floor(jml5rb)
            sisa5rb = sisa10rb - (5000 * Int(jml5rb))
            jml2rb = sisa5rb / 2000
            txt2rb.Text = Math.Floor(jml2rb)
            sisa2rb = sisa5rb - (2000 * Int(jml2rb))
            jml1rb = sisa2rb / 1000
            txt1rb.Text = Math.Floor(jml1rb)
            sisa1rb = sisa2rb - (1000 * Int(jml1rb))
            jml500 = sisa1rb / 500
            txt5rts.Text = Math.Floor(jml500)
            sisa500 = sisa1rb - (500 * Int(jml500))
        End If
    End Sub
    Private Sub Tbtotal_TextChanged(sender As Object, e As EventArgs) Handles tbtotal.TextChanged
        Call pecahanuang()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim m As String
        m = MsgBox("Apakah anda sudah menyimpan untuk struk ?", vbYesNo, "Medica")
        If m = vbYes Then
            tbsubtotal.Text = ""
            tbjumlahbeli.Text = ""
            tbkode__obat.Text = ""
            tbnama__obat.Text = ""
            tbharga.Text = ""
            tbstoktransaksi.Text = ""
            tbtotal.Text = ""
            tbbayar.Text = ""
            tbkembalian.Text = ""
            tbkodetransaksi.Text = ""
            tbstoktransaksi.Text = ""
            tbjumlahbeli.Enabled = False
            Call panggildatatransaksi()
            Call kodeautotransaksi()
        End If
    End Sub

    Private Sub Btncaritransaksi_Click(sender As Object, e As EventArgs) Handles btncaritransaksi.Click
        tbpencariantransaksi.Show()
    End Sub

    Private Sub tbjumlahbeli_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbjumlahbeli.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btnlaporanpasok.Click
        AxCrystalReport1.ReportFileName = "laporanpemasokan.rpt"
        AxCrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized
        AxCrystalReport1.RetrieveDataFiles()
        AxCrystalReport1.Action = 1
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        AxCrystalReport2.ReportFileName = "laporantransaksi.rpt"
        AxCrystalReport2.WindowState = Crystal.WindowStateConstants.crptMaximized
        AxCrystalReport2.RetrieveDataFiles()
        AxCrystalReport2.Action = 1
    End Sub

    Private Sub Btnstruk_Click(sender As Object, e As EventArgs) Handles btnstruk.Click
        Dim m As String
        If tbtotal.Text = "" Or tbbayar.Text = "" Or tbkembalian.Text = "" Or dgvtransaksi.CurrentCell Is Nothing Then
            MsgBox("Tidak dapat mencetak jika data masih belum terisi lengkap")
        Else
            m = MsgBox("Ingin membuat struk ?", vbYesNo, "Medica")
            If m = vbYes Then
                sqlnyastruk = "UPDATE tb_transaksi set bayar='" & tbbayar.Text & "',kembalian='" & tbkembalian.Text & "' where kd_transaksi='" & tbkodetransaksi.Text & "'"
                Call jalanstruk()
                MsgBox("Berhasil memuat")
                AxCrystalReport3.SelectionFormula = "totext({qw_transaksi.kd_transaksi})='" & tbkodetransaksi.Text & "'"
                AxCrystalReport3.ReportFileName = "laporannota.rpt"
                AxCrystalReport3.WindowState = Crystal.WindowStateConstants.crptMaximized
                AxCrystalReport3.RetrieveDataFiles()
                AxCrystalReport3.Action = 1
            End If
        End If
    End Sub
End Class