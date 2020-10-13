Imports System.Data.Odbc
Public Class returpenjualan
    Dim baris As Integer = 0
    Private Sub returpenjualan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Call Koneksi()
        Call mati()
        Call TampilGrid()
        Call kolomlist()
    End Sub
    Sub keluar()
        Me.Close()
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        DateTimePicker1.Value = Now
    End Sub
    Sub mati()
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        DateTimePicker1.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
    End Sub
    Sub hidup()
        TextBox10.Enabled = True
        TextBox9.Enabled = True
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
    End Sub
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select Kode_Transaksi,Tanggal,Keterangan,Kredit as Total_Penjualan From tb_bukupenjualan where keterangan IN ('Penjualan Tunai','Penjualan Kredit')", CONN)
        DS = New DataSet
        DA.Fill(DS, "beli")
        DataGridView1.DataSource = DS.Tables("beli")
        DataGridView1.ReadOnly = True
    End Sub
    Sub tampilgrid2()
        DA = New OdbcDataAdapter("select Kode_Transaksi,Kode_Barang,Nama_Barang,Satuan,dijual as Jumlah_barang from tb_bukupersediaanbarang where Kode_Transaksi='" & TextBox2.Text & "'", CONN)
        DS = New DataSet
        DA.Fill(DS, "barang")
        DataGridView2.DataSource = DS.Tables("barang")
        DataGridView2.ReadOnly = True
    End Sub
    Sub TampilGridhutang()
        DA = New OdbcDataAdapter("select * From tb_bukupiutang order by ID desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "piutang")
        DataGridView3.DataSource = DS.Tables("piutang")
        DataGridView3.ReadOnly = True
    End Sub
    Sub TampilGridkas()
        DA = New OdbcDataAdapter("select * From tb_bukukas order by ID desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "kas")
        DataGridView3.DataSource = DS.Tables("kas")
        DataGridView3.ReadOnly = True
    End Sub
    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        For f As Integer = 0 To DataGridView1.RowCount - 1
            If IsDBNull(DataGridView1.CurrentRow.Cells(0).Value) Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(1).Value) Then
                DateTimePicker1.Value = Now
            Else
                DateTimePicker1.Value = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(2).Value) Then
                TextBox3.Text = ""
            Else
                TextBox3.Text = DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(3).Value) Then
                TextBox4.Text = ""
            Else
                TextBox4.Text = DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
            End If
        Next
        DataGridView2.DataSource = Nothing
        Call tampilgrid2()
    End Sub

    Private Sub DataGridView2_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseDoubleClick
        For f As Integer = 0 To DataGridView1.RowCount - 1
            If IsDBNull(DataGridView1.CurrentRow.Cells(1).Value) Then
                TextBox5.Text = ""
            Else
                TextBox5.Text = DataGridView2.Item(1, DataGridView2.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(2).Value) Then
                TextBox6.Text = ""
            Else
                TextBox6.Text = DataGridView2.Item(2, DataGridView2.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(3).Value) Then
                TextBox7.Text = ""
            Else
                TextBox7.Text = DataGridView2.Item(3, DataGridView2.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(4).Value) Then
                TextBox8.Text = ""
            Else
                TextBox8.Text = DataGridView2.Item(4, DataGridView2.CurrentRow.Index).Value
            End If
        Next
    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged
        If Not IsNumeric(TextBox9.Text) Then
            TextBox9.Text = 0
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged
        If Not IsNumeric(TextBox10.Text) Then
            TextBox10.Text = 0
        End If
        TextBox11.Text = CDbl(TextBox10.Text) * CDbl(TextBox9.Text)
    End Sub
    Sub kolomlist()
        ListView1.View = View.Details
        ListView1.GridLines = True
        ListView1.Columns.Add("No", 30)
        ListView1.Columns.Add("Kode Barang", 80)
        ListView1.Columns.Add("Nama Barang", 100)
        ListView1.Columns.Add("Satuan", 80)
        ListView1.Columns.Add("Harga", 80)
        ListView1.Columns.Add("Jumlah", 80)
        ListView1.Columns.Add("Total", 80)
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If TextBox5.Text = "" Then
            MsgBox("Masukkan Barang Yang Akan Diretur", vbCritical + vbOKOnly, "Peringatan")
        ElseIf TextBox10.Text > TextBox8.Text Then
            MsgBox("Jumlah Retur Melebihi Pembelian Awal", vbCritical + vbOKOnly, "Peringatan")
        Else
            With ListView1
                .Items.Add(baris + 1)
                .Items(baris).SubItems.Add(TextBox5.Text)
                .Items(baris).SubItems.Add(TextBox6.Text)
                .Items(baris).SubItems.Add(TextBox7.Text)
                .Items(baris).SubItems.Add(TextBox9.Text)
                .Items(baris).SubItems.Add(TextBox10.Text)
                .Items(baris).SubItems.Add(TextBox11.Text)
            End With
            Label14.Text += CInt(TextBox11.Text)
            baris += 1
        End If
    End Sub
    Sub kodeotomatis()
        CMD = New OdbcCommand("select * from tb_retur order by Kode_Retur desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox1.Text = "RB" + "0001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Retur").ToString, 4, 3)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "RB000" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 2 Then
                TextBox1.Text = "RB00" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 3 Then
                TextBox1.Text = "RB0" & TextBox1.Text & ""
            End If
        End If
    End Sub
    Sub kodeotomatispenjualan()
        CMD = New OdbcCommand("select * from tb_bukupenjualan order by Kode_Transaksi desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            Label15.Text = "TJ" + "0001"
        Else
            Label15.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Transaksi").ToString, 4, 3)) + 1
            If Len(Label15.Text) = 1 Then
                Label15.Text = "TJ000" & Label15.Text & ""
            ElseIf Len(Label15.Text) = 2 Then
                Label15.Text = "TJ00" & Label15.Text & ""
            ElseIf Len(Label15.Text) = 3 Then
                Label15.Text = "TJ0" & Label15.Text & ""
            End If
        End If
    End Sub
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim tanggal As Integer
        tanggal = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        If DateTimePicker1.Value.Day = tanggal Then
            If TextBox3.Text = "Penjualan Kredit" Then

                DataGridView3.DataSource = Nothing
                Call TampilGridhutang()
                If DataGridView3.RowCount = 0 Then
                    Dim simpan7 As String = "insert into tb_bukupiutang values ('','" & TextBox1.Text & "','" & TextBox2.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Retur Barang','0','" & Label14.Text & "','" & Label14.Text & "','LUNAS')"
                    CMD = New OdbcCommand(simpan7, CONN)
                    CMD.ExecuteNonQuery()
                Else
                    Dim hitung, sisa As Double
                    sisa = CDbl(Label14.Text)
                    hitung = CDbl(DataGridView3.Item(7, 0).Value) - sisa
                    Dim simpan7 As String = "insert into tb_bukupiutang values ('','" & TextBox1.Text & "','" & TextBox2.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Retur Barang','0','" & sisa & "','" & hitung & "','LUNAS')"
                    CMD = New OdbcCommand(simpan7, CONN)
                    CMD.ExecuteNonQuery()
                    CMD = New OdbcCommand("select id,debet from tb_bukupiutang where Kode_Transaksi_Penjualan = '" & TextBox2.Text & "' AND Status='BELUM LUNAS' order by id desc", CONN)
                    RD = CMD.ExecuteReader
                    RD.Read()
                    Dim updatedebet As Double = RD.Item(1) - sisa
                    Dim updatestatus As String = "update tb_bukupiutang set Status='LUNAS' where kode_Transaksi_Penjualan = '" & TextBox2.Text & "' AND Status='BELUM LUNAS' order by id desc"
                    CMD = New OdbcCommand(updatestatus, CONN)
                    CMD.ExecuteNonQuery()
                    Dim simpan3 As String = "insert into tb_bukupiutang values ('','-','" & TextBox2.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Kredit','" & updatedebet & "','0','" & hitung & "','BELUM LUNAS')"
                    CMD = New OdbcCommand(simpan3, CONN)
                    CMD.ExecuteNonQuery()
                End If

            ElseIf TextBox3.Text = "Penjualan Tunai" Then
                DataGridView3.DataSource = Nothing
                Call TampilGridkas()
                If DataGridView3.RowCount = 0 Then
                    Dim simpan2 As String = "insert into tb_bukukas values ('','" & TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Retur Barang','0','" & Label14.Text & "','" & Label14.Text & "')"
                    CMD = New OdbcCommand(simpan2, CONN)
                    CMD.ExecuteNonQuery()
                Else
                    Dim hitung As Integer
                    hitung = CDbl(DataGridView3.Item(6, 0).Value) - CDbl(Label14.Text)
                    Dim simpan3 As String = "insert into tb_bukukas values ('','" & TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Retur Barang','0','" & Label14.Text & "','" & hitung & "')"
                    CMD = New OdbcCommand(simpan3, CONN)
                    CMD.ExecuteNonQuery()
                End If
            End If
            CMD = New OdbcCommand("select * from tb_bukupenjualan order by kode_transaksi desc", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            Dim saldopenjualan As Double = RD.Item(5) - CDbl(Label14.Text)
            kodeotomatispenjualan()
            Dim simpanreturpenjualan As String = "insert into tb_bukupenjualan values ('" & Label15.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Retur Barang','" & Label14.Text & "','0','" & saldopenjualan & "')"
            CMD = New OdbcCommand(simpanreturpenjualan, CONN)
            CMD.ExecuteNonQuery()

            'Dim simpan1 As String = "insert into tb_retur values ('" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','" & Label14.Text & "','0')"
            'CMD = New OdbcCommand(simpan1, CONN)
            'CMD.ExecuteNonQuery()

            With ListView1
                For i = 1 To .Items.Count
                    'Dim simpan As String = "insert into tb_bukupersediaanbarang values ('','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & .Items(i - 1).SubItems(1).Text & "','" & .Items(i - 1).SubItems(2).Text & "','" & .Items(i - 1).SubItems(3).Text & "','0','0','" & ListView1.Items(i - 1).SubItems(5).Text & "')"
                    'CMD = New OdbcCommand(simpan, CONN)
                    'CMD.ExecuteNonQuery()

                    CMD = New OdbcCommand("select * FROM tb_returrinci where Kode_Barang= '" & ListView1.Items(i - 1).SubItems(1).Text & "'", CONN)
                    RD = CMD.ExecuteReader
                    RD.Read()
                    If Not RD.HasRows Then
                        Dim simpanbarang As String = "insert into tb_returrinci values ('" & ListView1.Items(i - 1).SubItems(1).Text & "','" & ListView1.Items(i - 1).SubItems(2).Text & "','" & ListView1.Items(i - 1).SubItems(3).Text & "','" & ListView1.Items(i - 1).SubItems(5).Text & "')"
                        CMD = New OdbcCommand(simpanbarang, CONN)
                        CMD.ExecuteNonQuery()
                    Else
                        Dim stok As Integer = CDbl(RD.Item(3)) + CDbl(ListView1.Items(i - 1).SubItems(5).Text)
                        Dim edit As String = "update tb_returrinci set jumlah_retur='" & stok & "' where Kode_Barang='" & ListView1.Items(i - 1).SubItems(1).Text & "'"
                        CMD = New OdbcCommand(edit, CONN)
                        CMD.ExecuteNonQuery()
                    End If
                Next
                MsgBox("Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi")
                .Items.Clear()
                baris = 0
                Label14.Text = 0
                Call bersih()
                Call mati()
            End With
        Else
            If TextBox3.Text = "Penjualan Kredit" Then

                DataGridView3.DataSource = Nothing
                Call TampilGridhutang()
                If DataGridView3.RowCount = 0 Then
                    Dim simpan7 As String = "insert into tb_bukupiutang values ('','" & TextBox1.Text & "','" & TextBox2.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Retur Barang','0','" & Label14.Text & "','" & Label14.Text & "','LUNAS')"
                    CMD = New OdbcCommand(simpan7, CONN)
                    CMD.ExecuteNonQuery()
                Else
                    Dim hitung, sisa As Double
                    sisa = CDbl(Label14.Text)
                    hitung = CDbl(DataGridView3.Item(7, 0).Value) - sisa
                    Dim simpan7 As String = "insert into tb_bukupiutang values ('','" & TextBox1.Text & "','" & TextBox2.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Retur Barang','0','" & sisa & "','" & hitung & "','LUNAS')"
                    CMD = New OdbcCommand(simpan7, CONN)
                    CMD.ExecuteNonQuery()
                    CMD = New OdbcCommand("select id,debet from tb_bukupiutang where Kode_Transaksi_Penjualan = '" & TextBox2.Text & "' AND Status='BELUM LUNAS' order by id desc", CONN)
                    RD = CMD.ExecuteReader
                    RD.Read()
                    Dim updatedebet As Double = RD.Item(1) - sisa
                    Dim updatestatus As String = "update tb_bukupiutang set Status='LUNAS' where kode_Transaksi_Penjualan = '" & TextBox2.Text & "' AND Status='BELUM LUNAS' order by id desc"
                    CMD = New OdbcCommand(updatestatus, CONN)
                    CMD.ExecuteNonQuery()
                    Dim simpan3 As String = "insert into tb_bukupiutang values ('','-','" & TextBox2.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Kredit','" & updatedebet & "','0','" & hitung & "','BELUM LUNAS')"
                    CMD = New OdbcCommand(simpan3, CONN)
                    CMD.ExecuteNonQuery()
                End If

            ElseIf TextBox3.Text = "Penjualan Tunai" Then
                DataGridView3.DataSource = Nothing
                Call TampilGridkas()
                If DataGridView3.RowCount = 0 Then
                    Dim simpan2 As String = "insert into tb_bukukas values ('','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Retur Barang','0','" & Label14.Text & "','" & Label14.Text & "')"
                    CMD = New OdbcCommand(simpan2, CONN)
                    CMD.ExecuteNonQuery()
                Else
                    Dim hitung As Integer
                    hitung = CDbl(DataGridView3.Item(6, 0).Value) - CDbl(Label14.Text)
                    Dim simpan3 As String = "insert into tb_bukukas values ('','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Retur Barang','0','" & Label14.Text & "','" & hitung & "')"
                    CMD = New OdbcCommand(simpan3, CONN)
                    CMD.ExecuteNonQuery()
                End If
            End If
            CMD = New OdbcCommand("select * from tb_bukupenjualan order by kode_transaksi desc", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            Dim saldopenjualan As Double = RD.Item(5) - CDbl(Label14.Text)
            kodeotomatispenjualan()
            Dim simpanreturpenjualan As String = "insert into tb_bukupenjualan values ('" & Label15.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Retur Barang','" & Label14.Text & "','0','" & saldopenjualan & "')"
            CMD = New OdbcCommand(simpanreturpenjualan, CONN)
            CMD.ExecuteNonQuery()

            'Dim simpan1 As String = "insert into tb_retur values ('" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','" & Label14.Text & "','0')"
            'CMD = New OdbcCommand(simpan1, CONN)
            'CMD.ExecuteNonQuery()

            With ListView1
                For i = 1 To .Items.Count
                    'Dim simpan As String = "insert into tb_bukupersediaanbarang values ('','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & .Items(i - 1).SubItems(1).Text & "','" & .Items(i - 1).SubItems(2).Text & "','" & .Items(i - 1).SubItems(3).Text & "','0','0','" & ListView1.Items(i - 1).SubItems(5).Text & "')"
                    'CMD = New OdbcCommand(simpan, CONN)
                    'CMD.ExecuteNonQuery()

                    CMD = New OdbcCommand("select * FROM tb_returrinci where Kode_Barang= '" & ListView1.Items(i - 1).SubItems(1).Text & "'", CONN)
                    RD = CMD.ExecuteReader
                    RD.Read()
                    If Not RD.HasRows Then
                        Dim simpanbarang As String = "insert into tb_returrinci values ('" & ListView1.Items(i - 1).SubItems(1).Text & "','" & ListView1.Items(i - 1).SubItems(2).Text & "','" & ListView1.Items(i - 1).SubItems(3).Text & "','" & ListView1.Items(i - 1).SubItems(5).Text & "')"
                        CMD = New OdbcCommand(simpanbarang, CONN)
                        CMD.ExecuteNonQuery()
                    Else
                        Dim stok As Integer = CDbl(RD.Item(3)) + CDbl(ListView1.Items(i - 1).SubItems(5).Text)
                        Dim edit As String = "update tb_returrinci set jumlah_retur='" & stok & "' where Kode_Barang='" & ListView1.Items(i - 1).SubItems(1).Text & "'"
                        CMD = New OdbcCommand(edit, CONN)
                        CMD.ExecuteNonQuery()
                    End If
                Next
                MsgBox("Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi")
                .Items.Clear()
                baris = 0
                Label14.Text = 0
                Call bersih()
                Call mati()
            End With
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MsgBox("Pilih Barang yang Akan Diretur", vbCritical + vbOKOnly, "Peringatan")
        Else
            Call kodeotomatis()
            hidup()
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Call bersih()
        Call mati()
        Button1.Enabled = True
    End Sub
End Class