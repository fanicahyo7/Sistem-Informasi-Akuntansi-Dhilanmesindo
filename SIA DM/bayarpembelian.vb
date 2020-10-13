Imports System.Data.Odbc
Public Class bayarpembelian

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        With transaksipembelian
            With .ListView1
                If CDbl(TextBox1.Text) < CDbl(Label3.Text) Then
                    transaksipembelian.DataGridView1.DataSource = Nothing
                    Call transaksipembelian.TampilGridhutang()
                    If transaksipembelian.DataGridView1.RowCount = 0 Then
                        Dim simpan7 As String = "insert into tb_bukuhutang values ('','-','" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembelian Kredit','0','" & TextBox1.Text & "','" & TextBox1.Text & "','BELUM LUNAS')"
                        CMD = New OdbcCommand(simpan7, CONN)
                        CMD.ExecuteNonQuery()
                    Else
                        Dim hitung, sisa As Double
                        sisa = CDbl(Label3.Text) - CDbl(TextBox1.Text)
                        hitung = CDbl(transaksipembelian.DataGridView1.Item(7, 0).Value) + sisa
                        Dim simpan7 As String = "insert into tb_bukuhutang values ('','-','" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembelian Kredit','0','" & sisa & "','" & hitung & "','BELUM LUNAS')"
                        CMD = New OdbcCommand(simpan7, CONN)
                        CMD.ExecuteNonQuery()
                    End If

                    transaksipembelian.DataGridView1.DataSource = Nothing
                    Call transaksipembelian.TampilGridpembelian()
                    If transaksipembelian.DataGridView1.RowCount = 0 Then
                        Dim simpan4 As String = "insert into tb_bukupembelian values ('" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembelian Kredit','" & TextBox1.Text & "','0','" & TextBox1.Text & "')"
                        CMD = New OdbcCommand(simpan4, CONN)
                        CMD.ExecuteNonQuery()
                    Else
                        Dim hitunga As Integer
                        hitunga = CDbl(transaksipembelian.DataGridView1.Item(5, 0).Value) + CDbl(TextBox1.Text)
                        Dim simpan5 As String = "insert into tb_bukupembelian values ('" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembelian Kredit','" & TextBox1.Text & "','0','" & hitunga & "')"
                        CMD = New OdbcCommand(simpan5, CONN)
                        CMD.ExecuteNonQuery()
                    End If

                ElseIf CDbl(TextBox1.Text) >= CDbl(Label3.Text) Then
                    transaksipembelian.DataGridView1.DataSource = Nothing
                    Call transaksipembelian.TampilGridKas()
                    If transaksipembelian.DataGridView1.RowCount = 0 Then
                        Dim simpan2 As String = "insert into tb_bukukas values ('','" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembelian Tunai','0','" & Label3.Text & "','" & Label3.Text & "')"
                        CMD = New OdbcCommand(simpan2, CONN)
                        CMD.ExecuteNonQuery()
                    Else
                        Dim hitung As Integer
                        hitung = CDbl(transaksipembelian.DataGridView1.Item(6, 0).Value) - CDbl(Label3.Text)
                        Dim simpan3 As String = "insert into tb_bukukas values ('','" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembelian Tunai','0','" & Label3.Text & "','" & hitung & "')"
                        CMD = New OdbcCommand(simpan3, CONN)
                        CMD.ExecuteNonQuery()
                    End If

                    transaksipembelian.DataGridView1.DataSource = Nothing
                    Call transaksipembelian.TampilGridpembelian()
                    If transaksipembelian.DataGridView1.RowCount = 0 Then
                        Dim simpan4 As String = "insert into tb_bukupembelian values ('" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembelian Tunai','" & Label3.Text & "','0','" & Label3.Text & "')"
                        CMD = New OdbcCommand(simpan4, CONN)
                        CMD.ExecuteNonQuery()
                    Else
                        Dim hitunga As Integer
                        hitunga = CDbl(transaksipembelian.DataGridView1.Item(5, 0).Value) + CDbl(Label3.Text)
                        Dim simpan5 As String = "insert into tb_bukupembelian values ('" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembelian Tunai','" & Label3.Text & "','0','" & hitunga & "')"
                        CMD = New OdbcCommand(simpan5, CONN)
                        CMD.ExecuteNonQuery()
                    End If
                End If

                For i = 1 To .Items.Count
                    Dim simpan As String = "insert into tb_bukupersediaanbarang values ('','" & transaksipembelian.TextBox1.Text & "','" & Format(transaksipembelian.DateTimePicker1.Value, "yyyy-MM-dd") & "','" & transaksipembelian.ListView1.Items(i - 1).SubItems(1).Text & "','" & transaksipembelian.ListView1.Items(i - 1).SubItems(2).Text & "','" & transaksipembelian.ListView1.Items(i - 1).SubItems(3).Text & "','" & transaksipembelian.ListView1.Items(i - 1).SubItems(5).Text & "','0','0')"
                    CMD = New OdbcCommand(simpan, CONN)
                    CMD.ExecuteNonQuery()

                    CMD = New OdbcCommand("select * FROM tb_barang where Kode_Barang= '" & transaksipembelian.ListView1.Items(i - 1).SubItems(1).Text & "'", CONN)
                    RD = CMD.ExecuteReader
                    RD.Read()
                    Dim stok As Integer = CDbl(RD.Item(3)) + CDbl(transaksipembelian.ListView1.Items(i - 1).SubItems(5).Text)
                    Dim saldo As Double = CDbl(RD.Item(4)) + CDbl(transaksipembelian.ListView1.Items(i - 1).SubItems(6).Text)
                    Dim edit As String = "update tb_barang set stok='" & stok & "', saldo='" & saldo & "' where Kode_Barang='" & transaksipembelian.ListView1.Items(i - 1).SubItems(1).Text & "'"
                    CMD = New OdbcCommand(edit, CONN)
                    CMD.ExecuteNonQuery()
                Next
                MsgBox("Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi")
                .Items.Clear()
                transaksipembelian.baris = 0
                transaksipembelian.Label10.Text = 0
                Call transaksipembelian.bersih()
                Call transaksipembelian.isimati()
                transaksipembelian.TextBox4.Text = 0
                transaksipembelian.TextBox5.Text = 0
                Me.Dispose()
            End With
        End With
    End Sub

    Private Sub bayarpembelian_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
    End Sub
End Class