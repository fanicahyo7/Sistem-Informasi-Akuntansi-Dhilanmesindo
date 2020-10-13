Imports System.Data.Odbc
Public Class Form1
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select Jumlah from tb_modal order by ID desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "jual")
        DataGridView1.DataSource = DS.Tables("jual")
        DataGridView1.ReadOnly = True
    End Sub
    Sub TampilGrid2()
        DA = New OdbcDataAdapter("select Tanggal from tb_neraca where Tanggal='" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'", CONN)
        DS = New DataSet
        DA.Fill(DS, "ner")
        DataGridView2.DataSource = DS.Tables("ner")
        DataGridView2.ReadOnly = True
    End Sub
    Private Sub Form1_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2
        Me.IsMdiContainer = True
        Call Koneksi()
        CMD = New OdbcCommand("select * from tb_modal order by id desc", CONN)
        RD = CMD.ExecuteReader()
        RD.Read()
        If Not RD.HasRows Then
            frmmodal.MdiParent = Me
            frmmodal.Show()
        End If
        DataGridView1.DataSource = Nothing
        TampilGrid()
        DateTimePicker1.Value = Now
        'TextBox1.Text = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))

        'If DateTimePicker1.Value.Day = TextBox1.Text Then
        '    TextBox2.Text = "KUY DIPROSES"
        '    Dim tanggal, bulan, tahun As String
        '    bulan = DateTimePicker1.Value.Month
        '    tahun = DateTimePicker1.Value.Year
        '    tanggal = 1
        '    DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal
        '    CMD = New OdbcCommand("select sum(kredit) FROM tb_bukukas where tanggal >= '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' AND tanggal <= tanggal <= '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'", CONN)
        '    RD = CMD.ExecuteReader
        '    RD.Read()
        '    TextBox2.Text = RD.Item(0)
        '    'Dim stok As Integer = CDbl(RD.Item(3)) - CDbl(transaksipenjualan.ListView1.Items(i - 1).SubItems(5).Text)
        '    'Dim saldo As Double = CDbl(RD.Item(4)) - CDbl(transaksipenjualan.ListView1.Items(i - 1).SubItems(6).Text)
        '    'Dim edit As String = "update tb_barang set stok='" & stok & "', saldo='" & saldo & "' where Kode_Barang='" & transaksipenjualan.ListView1.Items(i - 1).SubItems(1).Text & "'"
        '    'CMD = New OdbcCommand(edit, CONN)
        '    'CMD.ExecuteNonQuery()
        'Else
        '    TextBox2.Text = "KADIT"
        '    Dim tanggal, bulan, tahun As String
        '    bulan = DateTimePicker1.Value.Month
        '    tahun = DateTimePicker1.Value.Year
        '    tanggal = 1
        '    DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal
        '    CMD = New OdbcCommand("select sum(kredit) FROM tb_bukukas where tanggal between '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'", CONN)
        '    RD = CMD.ExecuteReader
        '    RD.Read()
        '    TextBox2.Text = RD.Item(0)
        'End If
        TampilGrid2()
        If DataGridView2.RowCount <= 1 Then
            Dim saldokas, saldohutang, modal, saldopenjualan, saldopiutang, saldopotongan As Double
            TextBox1.Text = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))

            If DateTimePicker1.Value.Day = TextBox1.Text Then
                TextBox2.Text = "KUY DIPROSES"
                Dim tanggal, bulan, tahun As String
                bulan = DateTimePicker1.Value.Month
                tahun = DateTimePicker1.Value.Year
                tanggal = 1
                DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal
                'CMD = New OdbcCommand("select sum(kredit) FROM tb_bukukas where tanggal between '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'", CONN)
                'RD = CMD.ExecuteReader
                'RD.Read()
                'TextBox2.Text = RD.Item(0)


                CMD = New OdbcCommand("select saldo from tb_bukukas order by id desc", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    saldokas = 0
                Else
                    saldokas = RD.Item(0)
                End If
                Dim simpanawal As String = "insert into tb_bukukas values ('','-','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Saldo Akhir Kas','" & saldokas & "','0','" & saldokas & "')"
                CMD = New OdbcCommand(simpanawal, CONN)
                CMD.ExecuteNonQuery()
                Dim simpanakhir As String = "insert into tb_bukukas values ('','-','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Saldo Awal Kas','" & saldokas & "','0','" & saldokas & "')"
                CMD = New OdbcCommand(simpanakhir, CONN)
                CMD.ExecuteNonQuery()

                CMD = New OdbcCommand("select saldo from tb_bukukas order by id desc", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    saldokas = 0
                Else
                    saldokas = RD.Item(0)
                End If
                Dim simpankas As String = "insert into tb_neraca values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Kas','" & saldokas & "','-','0')"
                CMD = New OdbcCommand(simpankas, CONN)
                CMD.ExecuteNonQuery()

                CMD = New OdbcCommand("select saldo from tb_bukupiutang where Status='BELUM LUNAS' order by id desc", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    saldopiutang = 0
                Else
                    saldopiutang = RD.Item(0)
                End If
                Dim simpanpiutang As String = "insert into tb_neraca values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Piutang','" & saldopiutang & "','-','0')"
                CMD = New OdbcCommand(simpanpiutang, CONN)
                CMD.ExecuteNonQuery()



                CMD = New OdbcCommand("select sum(Jumlah) from tb_potongan where tanggal between '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    saldopotongan = 0
                Else
                    saldopotongan = RD.Item(0)
                End If
                Dim simpanpotongan As String = "insert into tb_neraca values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Potongan','" & saldopotongan & "','-','0')"
                CMD = New OdbcCommand(simpanpotongan, CONN)
                CMD.ExecuteNonQuery()

                'CMD = New OdbcCommand("select sum(saldo) from tb_barang", CONN)
                'RD = CMD.ExecuteReader
                'RD.Read()
                'saldobarang = RD.Item(0)
                'Dim simpanpersediaan As String = "insert into tb_neraca values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Persediaan Barang','" & saldobarang & "','-','0')"
                'CMD = New OdbcCommand(simpanpersediaan, CONN)
                'CMD.ExecuteNonQuery()

                CMD = New OdbcCommand("select saldo from tb_bukuhutang where Status='BELUM LUNAS' order by kode_hutang desc", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    saldohutang = 0
                Else
                    saldohutang = RD.Item(0)
                End If
                Dim simpanhutang As String = "insert into tb_neraca values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','-','0','Hutang','" & saldohutang & "')"
                CMD = New OdbcCommand(simpanhutang, CONN)
                CMD.ExecuteNonQuery()

                CMD = New OdbcCommand("select saldo from tb_bukupenjualan order by kode_transaksi desc", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    saldopenjualan = 0
                Else
                    saldopenjualan = RD.Item(0)
                End If
                Dim simpanneracapenjualan As String = "insert into tb_neraca values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','-','0','Penjualan','" & saldopenjualan & "')"
                CMD = New OdbcCommand(simpanneracapenjualan, CONN)
                CMD.ExecuteNonQuery()

                CMD = New OdbcCommand("select jumlah from tb_modal order by id desc", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    modal = 0
                Else
                    modal = RD.Item(0)
                End If
                Dim simpankasneracamodal As String = "insert into tb_neraca values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','-','0','Modal','" & modal & "')"
                CMD = New OdbcCommand(simpankasneracamodal, CONN)
                CMD.ExecuteNonQuery()

                'Dim tanggal, bulan, tahun As String
                'bulan = DateTimePicker1.Value.Month
                'tahun = DateTimePicker1.Value.Year
                'tanggal = 1
                'DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal
                'CMD = New OdbcCommand("select sum(kredit) FROM tb_bukukas where tanggal between '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'", CONN)
                'RD = CMD.ExecuteReader
                'RD.Read()
                'TextBox2.Text = RD.Item(0)
                Dim penjualan, biaya As Double
                CMD = New OdbcCommand("select saldo from tb_bukupenjualan order by Kode_Transaksi desc", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    penjualan = 0
                Else
                    penjualan = RD.Item(0)
                End If
                CMD = New OdbcCommand("select sum(biaya) from tb_bukubiaya where tanggal between '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Not RD.HasRows Then
                    biaya = 0
                Else
                    biaya = RD.Item(0)
                End If
                Dim simpanlabarugi As String = "insert into tb_labarugi values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan','" & penjualan & "','Biaya','" & biaya & "')"
                CMD = New OdbcCommand(simpanlabarugi, CONN)
                CMD.ExecuteNonQuery()

                Dim simpanpotongan2 As String = "insert into tb_labarugi values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','-','0','Potongan','" & saldopotongan & "')"
                CMD = New OdbcCommand(simpanpotongan2, CONN)
                CMD.ExecuteNonQuery()

                Dim laba As Double = penjualan - biaya - saldopotongan
                Dim simpansaldolaba As String = "insert into tb_labarugi values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Laba','" & laba & "','-','0')"
                CMD = New OdbcCommand(simpansaldolaba, CONN)
                CMD.ExecuteNonQuery()

                Dim ket As String
                If laba < 0 Then
                    ket = "Rugi"
                ElseIf laba > 0 Then
                    ket = "Laba"
                Else
                    ket = "Laba"
                End If

                DataGridView1.DataSource = Nothing
                TampilGrid()
                Dim jumlah As Double
                jumlah = DataGridView1.Item(0, 0).Value
                Dim jml As Double
                jml = jumlah + laba
                Dim simpansaldomodal As String = "insert into tb_modal values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & ket & " " & CStr(laba) & "','" & jml & "')"
                CMD = New OdbcCommand(simpansaldomodal, CONN)
                CMD.ExecuteNonQuery()

                Dim saldokaslagi As Double = saldokas + laba
                Dim simpanlabakas As String = "insert into tb_bukukas values ('','-','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Laba','" & laba & "','0','" & saldokaslagi & "')"
                CMD = New OdbcCommand(simpanlabakas, CONN)
                CMD.ExecuteNonQuery()


                'kodeotomatispenjualan()
                'CMD = New OdbcCommand("select saldo from tb_bukupenjualan order by kode_transaksi desc", CONN)
                'RD = CMD.ExecuteReader
                'RD.Read()
                'Dim simpansaldopenjualan As String = "insert into tb_bukupenjualan values ('" & TextBox2.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Saldo','0','0','" & RD.Item(0) & "'"
                'CMD = New OdbcCommand(simpansaldopenjualan, CONN)
                'CMD.ExecuteNonQuery()
            Else
                TextBox2.Text = "KADIT"
            End If
        End If
    End Sub

    Private Sub PersediaanBarangToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PersediaanBarangToolStripMenuItem.Click
        formmasterBarang.MdiParent = Me
        formmasterBarang.Show()

        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukukas.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub PembelianToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)
        transaksipembelian.Show()
    End Sub

    Private Sub KasToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles KasToolStripMenuItem.Click
        frmbukukas.MdiParent = Me
        frmbukukas.Show()

        formmasterBarang.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub PembelianToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        frmbukupembelian.Show()
    End Sub

    Private Sub PenjualanToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PenjualanToolStripMenuItem.Click
        frmbukupenjualan.MdiParent = Me
        frmbukupenjualan.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub HutangToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HutangToolStripMenuItem.Click
        frmbukuhutang.MdiParent = Me
        frmbukuhutang.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub PiutangToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PiutangToolStripMenuItem.Click
        frmbukupiutang.MdiParent = Me
        frmbukupiutang.Show()

        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukukas.keluar()
        frmbukupenjualan.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub PenjualanToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles PenjualanToolStripMenuItem1.Click
        transaksipenjualan.MdiParent = Me
        transaksipenjualan.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub BiayaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BiayaToolStripMenuItem.Click
        transaksibiaya.MdiParent = Me
        transaksibiaya.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub BiayaToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles BiayaToolStripMenuItem1.Click
        frmbukubiaya.MdiParent = Me
        frmbukubiaya.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukuhutang.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub HutangToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles HutangToolStripMenuItem1.Click
        frmhutang.MdiParent = Me
        frmhutang.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub
    Sub kekuar()
        Close()
    End Sub
    Private Sub PiutangToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles PiutangToolStripMenuItem1.Click
        bayarpiutang.MdiParent = Me
        bayarpiutang.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub ReturPembelianToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        returpembelian.Show()
    End Sub

    Private Sub ReturPenjualanToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ReturPenjualanToolStripMenuItem.Click
        returpenjualan.MdiParent = Me
        returpenjualan.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        'TextBox1.Text = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        'TextBox1.Text = Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd")
    End Sub
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker2.ValueChanged

        Dim tanggal, bulan, tahun As String
        bulan = DateTimePicker2.Value.Month
        tahun = DateTimePicker2.Value.Year
        tanggal = 1

        DateTimePicker3.Value = tahun + "-" + bulan + "-" + tanggal
    End Sub

    Private Sub NeracaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NeracaToolStripMenuItem.Click
        neraca.MdiParent = Me
        neraca.Show()

        frmbukukas.keluar()
        bayarhutang.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukukas.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub BayarHutangToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BayarHutangToolStripMenuItem.Click
        bayarhutang.MdiParent = Me
        bayarhutang.Show()

        frmbukukas.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukukas.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub
    Sub kodeotomatispenjualan()
        CMD = New OdbcCommand("select * from tb_bukupenjualan order by Kode_Transaksi desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox2.Text = "TJ" + "0001"
        Else
            TextBox2.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Transaksi").ToString, 4, 3)) + 1
            If Len(TextBox2.Text) = 1 Then
                TextBox2.Text = "TJ000" & TextBox2.Text & ""
            ElseIf Len(TextBox2.Text) = 2 Then
                TextBox2.Text = "TJ00" & TextBox2.Text & ""
            ElseIf Len(TextBox2.Text) = 3 Then
                TextBox2.Text = "TJ0" & TextBox2.Text & ""
            End If
        End If
    End Sub

    Private Sub KeluarAplikasiToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles KeluarAplikasiToolStripMenuItem.Click
        End
    End Sub

    Private Sub LabaRugiToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LabaRugiToolStripMenuItem.Click
        frmlabarugi.MdiParent = Me
        frmlabarugi.Show()

        frmbukukas.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukukas.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmperubahanmodal.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub PerubahanModalToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PerubahanModalToolStripMenuItem.Click
        frmperubahanmodal.MdiParent = Me
        frmperubahanmodal.Show()

        frmbukukas.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukukas.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmlabarugi.keluar()
        frmgantipass.keluar()
    End Sub

    Private Sub GantiPasswordToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GantiPasswordToolStripMenuItem.Click
        frmgantipass.MdiParent = Me
        frmgantipass.Show()

        frmbukukas.keluar()
        bayarpenjualan.keluar()
        bayarpiutang.keluar()
        frmbukubiaya.keluar()
        frmbukuhutang.keluar()
        frmbukukas.keluar()
        frmbukupenjualan.keluar()
        frmbukupiutang.keluar()
        frmhutang.keluar()
        formmasterBarang.keluar()
        neraca.keluar()
        returpenjualan.keluar()
        transaksibiaya.keluar()
        transaksipenjualan.keluar()
        frmlabarugi.keluar()
        frmperubahanmodal.keluar()
    End Sub
End Class
