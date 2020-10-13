Imports System.Data.Odbc
Public Class bayarpenjualan
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim tanggal As Integer
        With transaksipenjualan
            tanggal = Hasil(Year(.DateTimePicker1.Text), Month(.DateTimePicker1.Text))
            If .DateTimePicker1.Value.Day = tanggal Then
                With .ListView1
                    If CDbl(TextBox1.Text) < CDbl(Label3.Text) Then
                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridpiutang()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim sisa As Double
                            sisa = CDbl(Label3.Text) - CDbl(TextBox1.Text)
                            Dim simpan7 As String = "insert into tb_bukupiutang values ('','-','" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Kredit','" & sisa & "','0','" & sisa & "','BELUM LUNAS')"
                            CMD = New OdbcCommand(simpan7, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitung, sisa As Double
                            sisa = CDbl(Label3.Text) - CDbl(TextBox1.Text)
                            hitung = CDbl(transaksipenjualan.DataGridView1.Item(7, 0).Value) + sisa
                            Dim simpan7 As String = "insert into tb_bukupiutang values ('','-','" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Kredit','" & sisa & "','0','" & hitung & "','BELUM LUNAS')"
                            CMD = New OdbcCommand(simpan7, CONN)
                            CMD.ExecuteNonQuery()
                        End If

                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridpenjualan()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim simpan4 As String = "insert into tb_bukupenjualan values ('" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Kredit','0','" & Label3.Text & "','" & Label3.Text & "')"
                            CMD = New OdbcCommand(simpan4, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitunga As Integer
                            hitunga = CDbl(transaksipenjualan.DataGridView1.Item(5, 0).Value) + CDbl(Label3.Text)
                            Dim simpan5 As String = "insert into tb_bukupenjualan values ('" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Kredit','0','" & Label3.Text & "','" & hitunga & "')"
                            CMD = New OdbcCommand(simpan5, CONN)
                            CMD.ExecuteNonQuery()
                        End If

                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridKas()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim simpan2 As String = "insert into tb_bukukas values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Kredit','" & TextBox1.Text & "','0','" & TextBox1.Text & "')"
                            CMD = New OdbcCommand(simpan2, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitung As Integer
                            hitung = CDbl(transaksipenjualan.DataGridView1.Item(6, 0).Value) + CDbl(TextBox1.Text)
                            Dim simpan3 As String = "insert into tb_bukukas values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Kredit','" & TextBox1.Text & "','0','" & hitung & "')"
                            CMD = New OdbcCommand(simpan3, CONN)
                            CMD.ExecuteNonQuery()
                        End If

                    ElseIf CDbl(TextBox1.Text) >= CDbl(Label3.Text) Then
                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridKas()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim simpan2 As String = "insert into tb_bukukas values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Tunai','" & Label3.Text & "','0','" & Label3.Text & "')"
                            CMD = New OdbcCommand(simpan2, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitung As Integer
                            hitung = CDbl(transaksipenjualan.DataGridView1.Item(6, 0).Value) + CDbl(Label3.Text)
                            Dim simpan3 As String = "insert into tb_bukukas values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Tunai','" & Label3.Text & "','0','" & hitung & "')"
                            CMD = New OdbcCommand(simpan3, CONN)
                            CMD.ExecuteNonQuery()
                        End If

                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridpenjualan()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim simpan4 As String = "insert into tb_bukupenjualan values ('" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Tunai','0','" & Label3.Text & "','" & Label3.Text & "')"
                            CMD = New OdbcCommand(simpan4, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitunga As Integer
                            hitunga = CDbl(transaksipenjualan.DataGridView1.Item(5, 0).Value) + CDbl(Label3.Text)
                            Dim simpan5 As String = "insert into tb_bukupenjualan values ('" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','Penjualan Tunai','0','" & Label3.Text & "','" & hitunga & "')"
                            CMD = New OdbcCommand(simpan5, CONN)
                            CMD.ExecuteNonQuery()
                        End If
                    End If

                    For i = 1 To .Items.Count
                        Dim simpan As String = "insert into tb_bukupersediaanbarang values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, transaksipenjualan.DateTimePicker1.Value), "yyyy-MM-dd") & "','" & transaksipenjualan.ListView1.Items(i - 1).SubItems(1).Text & "','" & transaksipenjualan.ListView1.Items(i - 1).SubItems(2).Text & "','" & transaksipenjualan.ListView1.Items(i - 1).SubItems(3).Text & "','0','" & transaksipenjualan.ListView1.Items(i - 1).SubItems(5).Text & "','0')"
                        CMD = New OdbcCommand(simpan, CONN)
                        CMD.ExecuteNonQuery()

                        CMD = New OdbcCommand("select * FROM tb_barang where Kode_Barang= '" & transaksipenjualan.ListView1.Items(i - 1).SubItems(1).Text & "'", CONN)
                        RD = CMD.ExecuteReader
                        RD.Read()
                        Dim stok As Integer = CDbl(RD.Item(3)) - CDbl(transaksipenjualan.ListView1.Items(i - 1).SubItems(5).Text)
                        'Dim saldo As Double = CDbl(RD.Item(4)) - CDbl(transaksipenjualan.ListView1.Items(i - 1).SubItems(6).Text)
                        Dim edit As String = "update tb_barang set stok='" & stok & "' where Kode_Barang='" & transaksipenjualan.ListView1.Items(i - 1).SubItems(1).Text & "'"
                        CMD = New OdbcCommand(edit, CONN)
                        CMD.ExecuteNonQuery()
                    Next
                    MsgBox("Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi")
                    .Items.Clear()
                    transaksipenjualan.baris = 0
                    transaksipenjualan.Label10.Text = 0
                    Call transaksipenjualan.bersih()
                    Call transaksipenjualan.isimati()
                    transaksipenjualan.TextBox4.Text = 0
                    transaksipenjualan.TextBox5.Text = 0
                    Me.Dispose()
                End With
            Else
                With .ListView1
                    If CDbl(TextBox1.Text) < CDbl(Label3.Text) Then
                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridpiutang()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim sisa As Double
                            sisa = CDbl(Label3.Text) - CDbl(TextBox1.Text)
                            Dim simpan7 As String = "insert into tb_bukupiutang values ('','-','" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Kredit','" & sisa & "','0','" & sisa & "','BELUM LUNAS')"
                            CMD = New OdbcCommand(simpan7, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitung, sisa As Double
                            sisa = CDbl(Label3.Text) - CDbl(TextBox1.Text)
                            hitung = CDbl(transaksipenjualan.DataGridView1.Item(7, 0).Value) + sisa
                            Dim simpan7 As String = "insert into tb_bukupiutang values ('','-','" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Kredit','" & sisa & "','0','" & hitung & "','BELUM LUNAS')"
                            CMD = New OdbcCommand(simpan7, CONN)
                            CMD.ExecuteNonQuery()
                        End If

                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridpenjualan()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim simpan4 As String = "insert into tb_bukupenjualan values ('" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Kredit','0','" & Label3.Text & "','" & Label3.Text & "')"
                            CMD = New OdbcCommand(simpan4, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitunga As Integer
                            hitunga = CDbl(transaksipenjualan.DataGridView1.Item(5, 0).Value) + CDbl(Label3.Text)
                            Dim simpan5 As String = "insert into tb_bukupenjualan values ('" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Kredit','0','" & Label3.Text & "','" & hitunga & "')"
                            CMD = New OdbcCommand(simpan5, CONN)
                            CMD.ExecuteNonQuery()
                        End If

                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridKas()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim simpan2 As String = "insert into tb_bukukas values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Kredit','" & TextBox1.Text & "','0','" & TextBox1.Text & "')"
                            CMD = New OdbcCommand(simpan2, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitung As Integer
                            hitung = CDbl(transaksipenjualan.DataGridView1.Item(6, 0).Value) + CDbl(TextBox1.Text)
                            Dim simpan3 As String = "insert into tb_bukukas values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Kredit','" & TextBox1.Text & "','0','" & hitung & "')"
                            CMD = New OdbcCommand(simpan3, CONN)
                            CMD.ExecuteNonQuery()
                        End If

                    ElseIf CDbl(TextBox1.Text) >= CDbl(Label3.Text) Then
                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridKas()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim simpan2 As String = "insert into tb_bukukas values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Tunai','" & Label3.Text & "','0','" & Label3.Text & "')"
                            CMD = New OdbcCommand(simpan2, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitung As Integer
                            hitung = CDbl(transaksipenjualan.DataGridView1.Item(6, 0).Value) + CDbl(Label3.Text)
                            Dim simpan3 As String = "insert into tb_bukukas values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Tunai','" & Label3.Text & "','0','" & hitung & "')"
                            CMD = New OdbcCommand(simpan3, CONN)
                            CMD.ExecuteNonQuery()
                        End If

                        transaksipenjualan.DataGridView1.DataSource = Nothing
                        Call transaksipenjualan.TampilGridpenjualan()
                        If transaksipenjualan.DataGridView1.RowCount = 0 Then
                            Dim simpan4 As String = "insert into tb_bukupenjualan values ('" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Tunai','0','" & Label3.Text & "','" & Label3.Text & "')"
                            CMD = New OdbcCommand(simpan4, CONN)
                            CMD.ExecuteNonQuery()
                        Else
                            Dim hitunga As Integer
                            hitunga = CDbl(transaksipenjualan.DataGridView1.Item(5, 0).Value) + CDbl(Label3.Text)
                            Dim simpan5 As String = "insert into tb_bukupenjualan values ('" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','Penjualan Tunai','0','" & Label3.Text & "','" & hitunga & "')"
                            CMD = New OdbcCommand(simpan5, CONN)
                            CMD.ExecuteNonQuery()
                        End If
                    End If

                    For i = 1 To .Items.Count
                        Dim simpan As String = "insert into tb_bukupersediaanbarang values ('','" & transaksipenjualan.TextBox1.Text & "','" & Format(transaksipenjualan.DateTimePicker1.Value, "yyyy-MM-dd") & "','" & transaksipenjualan.ListView1.Items(i - 1).SubItems(1).Text & "','" & transaksipenjualan.ListView1.Items(i - 1).SubItems(2).Text & "','" & transaksipenjualan.ListView1.Items(i - 1).SubItems(3).Text & "','0','" & transaksipenjualan.ListView1.Items(i - 1).SubItems(5).Text & "','0')"
                        CMD = New OdbcCommand(simpan, CONN)
                        CMD.ExecuteNonQuery()

                        CMD = New OdbcCommand("select * FROM tb_barang where Kode_Barang= '" & transaksipenjualan.ListView1.Items(i - 1).SubItems(1).Text & "'", CONN)
                        RD = CMD.ExecuteReader
                        RD.Read()
                        Dim stok As Integer = CDbl(RD.Item(3)) - CDbl(transaksipenjualan.ListView1.Items(i - 1).SubItems(5).Text)
                        'Dim saldo As Double = CDbl(RD.Item(4)) - CDbl(transaksipenjualan.ListView1.Items(i - 1).SubItems(6).Text)
                        Dim edit As String = "update tb_barang set stok='" & stok & "' where Kode_Barang='" & transaksipenjualan.ListView1.Items(i - 1).SubItems(1).Text & "'"
                        CMD = New OdbcCommand(edit, CONN)
                        CMD.ExecuteNonQuery()
                    Next
                    MsgBox("Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi")
                    .Items.Clear()
                    transaksipenjualan.baris = 0
                    transaksipenjualan.Label10.Text = 0
                    Call transaksipenjualan.bersih()
                    Call transaksipenjualan.isimati()
                    transaksipenjualan.TextBox4.Text = 0
                    transaksipenjualan.TextBox5.Text = 0
                    transaksipenjualan.Button2.Enabled = True
                    Me.Dispose()
                End With
            End If
        End With
    End Sub

    Private Sub bayarpenjualan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
    End Sub
    Sub keluar()
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        If Not IsNumeric(TextBox1.Text) Then
            TextBox1.Text = 0
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dispose()
    End Sub
End Class