Imports System.Data.Odbc
Public Class formmasterBarang
    Private Sub PersediaanBarang_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Call segarkan()
    End Sub
    Sub keluar()
        Me.Close()
    End Sub
    Sub segarkan()
        Call Koneksi()
        Call TampilGrid()
        Call isimati()
        Call tombolhidup()
        Call bersih()
    End Sub
    Sub tombolhidup()
        Button1.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button2.Enabled = False
        Button1.Text = "Tambah"
        Button3.Text = "Ubah"
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
    Sub isimati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        Button2.Enabled = False
    End Sub
    Sub isihidup()
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
    End Sub
    Sub kodeotomatis()
        CMD = New OdbcCommand("select * from tb_barang order by Kode_Barang desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox1.Text = "BR" + "0001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Barang").ToString, 4, 3)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "BR000" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 2 Then
                TextBox1.Text = "BR00" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 3 Then
                TextBox1.Text = "BR0" & TextBox1.Text & ""
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "Tambah" Then
            Button1.Text = "Simpan"
            Button2.Enabled = True
            Button3.Enabled = False
            Button4.Enabled = False
            Call bersih()
            Call kodeotomatis()
            Call isihidup()
        ElseIf Button1.Text = "Simpan" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Data Belum Diisi Dengan Lengkap!", vbCritical + vbOKOnly, "Peringatan")
            Else
                Dim simpan As String = "insert into tb_barang values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil Disimpan!", vbInformation + vbOKOnly, "Informasi")
                Call segarkan()
            End If
        End If
    End Sub

    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From tb_barang", CONN)
        DS = New DataSet
        DA.Fill(DS, "barang")
        DataGridView1.DataSource = DS.Tables("barang")
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Call bersih()
        Call tombolhidup()
        Call isimati()
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        TextBox1.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        TextBox2.Text = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
        TextBox3.Text = DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
        TextBox4.Text = DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Pilih Data Yang Akan Diubah!", vbCritical + vbOKOnly, "Peringatan")
        ElseIf Button3.Text = "Ubah" Then
            Call isihidup()
            Button3.Text = "Simpan"
            Button2.Enabled = True
            Button1.Enabled = False
            Button4.Enabled = False
        ElseIf Button3.Text = "Simpan" Then
            Dim edit As String = "update tb_barang set Nama_Barang='" & TextBox2.Text & "',Satuan='" & TextBox3.Text & "',Stok='" & TextBox4.Text & "' where Kode_Barang='" & TextBox1.Text & "'"
            CMD = New OdbcCommand(edit, CONN)
            CMD.ExecuteNonQuery()
            MsgBox("Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi")
            Call segarkan()
        End If
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih Data Yang Akan Dihapus!", vbCritical + vbOKOnly, "Peringatan")
        Else
            If MsgBox("Anda Yakin Ingin Menghapus?", vbQuestion + vbYesNo) = vbYes Then
                Dim hapus As String = "delete From tb_barang where kode_barang='" & TextBox1.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi")
                Call segarkan()
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        'Dim rpt As New barang 'Report yang telah dibuat.
        'Try
        '    DA = New OdbcDataAdapter("select * From tb_barang", CONN)
        '    DA.Fill(DS, "barang")
        '    rpt.SetDataSource(DS)
        '    CrystalReportViewer1.ReportSource = rpt
        'Catch Excep As Exception
        '    MessageBox.Show(Excep.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        'ccccc.CrystalReportViewer1.SelectionFormula = ""
        reviewkas.CrystalReportViewer1.RefreshReport()
        ccccc.CrystalReportViewer1.Refresh()
        ccccc.Show()
    End Sub
End Class