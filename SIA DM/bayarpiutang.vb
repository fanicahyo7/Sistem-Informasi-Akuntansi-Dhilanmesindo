Imports System.Data.Odbc
Public Class bayarpiutang
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = 0
        DateTimePicker1.Value = Now
        Label1.Text = ""
        Label10.Text = 0
    End Sub
    Sub mati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        DateTimePicker1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
    End Sub
    Sub keluar()
        Me.Close()
    End Sub

    Private Sub bayarpiutang_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Call Koneksi()
        Call bersih()
        Call mati()
        Call tampilisi()
    End Sub
    Sub tampilisi()
        DA = New OdbcDataAdapter("select ID,Kode_Transaksi_Penjualan,Tanggal,Keterangan,Debet as Hutang from tb_bukupiutang where Status='BELUM LUNAS'", CONN)
        DS = New DataSet
        DA.Fill(DS, "utang")
        DataGridView1.DataSource = DS.Tables("utang")
        DataGridView1.ReadOnly = True
        DataGridView1.Refresh()
    End Sub
    Sub tampilisi2()
        DA = New OdbcDataAdapter("select * from tb_bukupiutang order by id desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "utang2")
        DataGridView2.DataSource = DS.Tables("utang2")
        DataGridView2.ReadOnly = True
        DataGridView2.Refresh()
    End Sub
    Sub tampilisikas()
        DA = New OdbcDataAdapter("select * from tb_bukukas order by id desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "kas")
        DataGridView2.DataSource = DS.Tables("kas")
        DataGridView2.ReadOnly = True
        DataGridView2.Refresh()
    End Sub
    Sub tampilisipotongan()
        DA = New OdbcDataAdapter("select * from tb_potongan order by id desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "ptg")
        DataGridView2.DataSource = DS.Tables("ptg")
        DataGridView2.ReadOnly = True
        DataGridView2.Refresh()
    End Sub
    Sub kodeotomatis()
        CMD = New OdbcCommand("select * from tb_bukupiutang order by id desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox5.Text = "BP" + "0001"
        Else
            TextBox5.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Transaksi_Penjualan").ToString, 4, 3)) + 1
            If Len(TextBox5.Text) = 1 Then
                TextBox5.Text = "BP000" & TextBox5.Text & ""
            ElseIf Len(TextBox5.Text) = 2 Then
                TextBox5.Text = "BP00" & TextBox5.Text & ""
            ElseIf Len(TextBox5.Text) = 3 Then
                TextBox5.Text = "BP0" & TextBox5.Text & ""
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        For f As Integer = 0 To DataGridView1.RowCount - 1
            If IsDBNull(DataGridView1.CurrentRow.Cells(0).Value) Then
                Label1.Text = ""
            Else
                Label1.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(1).Value) Then
                TextBox1.Text = ""
            Else
                TextBox1.Text = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(2).Value) Then
                DateTimePicker1.Value = Now
            Else
                DateTimePicker1.Value = DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(3).Value) Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(4).Value) Then
                TextBox3.Text = ""
                Label10.Text = "0"
            Else
                TextBox3.Text = DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value
                Label10.Text = DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value
            End If
        Next
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih Data yang Akan Dibayar", vbCritical + vbOKOnly, "Peringatan")
        Else
            TextBox4.Enabled = True
            TextBox6.Enabled = True
            Button1.Enabled = False
            Button2.Enabled = True
            Button3.Enabled = True
            Call kodeotomatis()
            TextBox6.Text = 0
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Call mati()
        Call bersih()
        Button1.Enabled = True
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged
        If Not IsNumeric(TextBox4.Text) Then
            TextBox4.Text = 0
        End If
    End Sub
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim tanggal As Integer
        tanggal = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        If DateTimePicker1.Value.Day = tanggal Then
            DataGridView2.DataSource = Nothing
            DateTimePicker1.Value = Now
            If CDbl(Label10.Text) = CDbl(TextBox4.Text) Then
                Call tampilisi2()
                Dim hitung As Double
                hitung = CDbl(DataGridView2.Item(7, 0).Value) - CDbl(TextBox3.Text)
                Dim simpan As String = "insert into tb_bukupiutang values ('','" & TextBox5.Text & "','" & TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Pembayaran Piutang','0','" & TextBox3.Text & "','" & hitung & "','LUNAS')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                Dim ubahstatus As String = "update tb_bukupiutang set Status='LUNAS' where ID='" & Label1.Text & "'"
                CMD = New OdbcCommand(ubahstatus, CONN)
                CMD.ExecuteNonQuery()

                DataGridView2.DataSource = Nothing
                Call tampilisikas()
                Dim hitungsaldo As Double
                hitungsaldo = CDbl(DataGridView2.Item(6, 0).Value) + CDbl(TextBox4.Text)
                Dim simpankas As String = "insert into tb_bukukas values ('','" & TextBox5.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Pembayaran Piutang','" & TextBox4.Text & "','0','" & hitungsaldo & "')"
                CMD = New OdbcCommand(simpankas, CONN)
                CMD.ExecuteNonQuery()

                Dim hitungsisa As Double
                hitungsisa = CDbl(TextBox3.Text) - CDbl(TextBox4.Text)
                Dim simpanpotongan As String
                simpanpotongan = "insert into tb_potongan values ('','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Potongan Penjualan '" & TextBox1.Text & "','" & hitungsisa & "')"
                CMD = New OdbcCommand(simpanpotongan, CONN)
                CMD.ExecuteNonQuery()

                MsgBox("Pembayaran Berhasil", vbInformation + vbOKOnly, "Informasi")
                Call bersih()
                Call mati()
                Button1.Enabled = True
                DataGridView1.DataSource = Nothing
                Call tampilisi()

            ElseIf CDbl(TextBox4.Text) < CDbl(Label10.Text) Then
                Call tampilisi2()
                Dim hitung As Double
                hitung = CDbl(DataGridView2.Item(7, 0).Value) - CDbl(TextBox3.Text)
                Dim simpan As String = "insert into tb_bukupiutang values ('','-','" & TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Pembayaran Piutang','0','" & TextBox3.Text & "','" & hitung & "','LUNAS')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                Dim ubahstatus As String = "update tb_bukupiutang set Status='LUNAS' where ID='" & Label1.Text & "'"
                CMD = New OdbcCommand(ubahstatus, CONN)
                CMD.ExecuteNonQuery()
                'simpan yang baru BELUM
                DataGridView2.DataSource = Nothing
                tampilisi2()
                Dim ss As Double = CDbl(Label10.Text) - CDbl(TextBox4.Text)
                Dim sld As Double = CDbl(DataGridView2.Item(7, 0).Value) + ss
                Dim simpanlagi As String = "insert into tb_bukupiutang values ('','-','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Pembayaran Piutang','" & ss & "','0','" & sld & "','BELUM LUNAS')"
                CMD = New OdbcCommand(simpanlagi, CONN)
                CMD.ExecuteNonQuery()

                DataGridView2.DataSource = Nothing
                Call tampilisikas()
                Dim hitungsaldo As Double
                hitungsaldo = CDbl(DataGridView2.Item(6, 0).Value) + CDbl(TextBox4.Text)
                Dim simpankas As String = "insert into tb_bukukas values ('','" & TextBox5.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','Pembayaran Piutang','" & TextBox4.Text & "','0','" & hitungsaldo & "')"
                CMD = New OdbcCommand(simpankas, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Pembayaran Berhasil", vbInformation + vbOKOnly, "Informasi")
                Call bersih()
                Call mati()
                Button1.Enabled = True
                DataGridView1.DataSource = Nothing
                Call tampilisi()
            End If
        Else
            DataGridView2.DataSource = Nothing
            DateTimePicker1.Value = Now
            If CDbl(Label10.Text) = CDbl(TextBox4.Text) Then
                Call tampilisi2()
                Dim hitung As Double
                hitung = CDbl(DataGridView2.Item(7, 0).Value) - CDbl(TextBox3.Text)
                Dim simpan As String = "insert into tb_bukupiutang values ('','" & TextBox5.Text & "','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembayaran Piutang','0','" & TextBox3.Text & "','" & hitung & "','LUNAS')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                Dim ubahstatus As String = "update tb_bukupiutang set Status='LUNAS' where ID='" & Label1.Text & "'"
                CMD = New OdbcCommand(ubahstatus, CONN)
                CMD.ExecuteNonQuery()

                DataGridView2.DataSource = Nothing
                Call tampilisikas()
                Dim hitungsaldo As Double
                hitungsaldo = CDbl(DataGridView2.Item(6, 0).Value) + CDbl(TextBox4.Text)
                Dim simpankas As String = "insert into tb_bukukas values ('','" & TextBox5.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembayaran Piutang','" & TextBox4.Text & "','0','" & hitungsaldo & "')"
                CMD = New OdbcCommand(simpankas, CONN)
                CMD.ExecuteNonQuery()

                Dim hitungsisa As Double
                hitungsisa = CDbl(TextBox3.Text) - CDbl(TextBox4.Text)
                Dim simpanpotongan As String
                simpanpotongan = "insert into tb_potongan values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Potongan Penjualan " & TextBox1.Text & "','" & hitungsisa & "')"
                CMD = New OdbcCommand(simpanpotongan, CONN)
                CMD.ExecuteNonQuery()

                MsgBox("Pembayaran Berhasil", vbInformation + vbOKOnly, "Informasi")
                Call bersih()
                Call mati()
                Button1.Enabled = True
                DataGridView1.DataSource = Nothing
                Call tampilisi()

            ElseIf CDbl(TextBox4.Text) < CDbl(Label10.Text) Then
                Call tampilisi2()
                Dim hitung As Double
                hitung = CDbl(DataGridView2.Item(7, 0).Value) - CDbl(TextBox3.Text)
                Dim simpan As String = "insert into tb_bukupiutang values ('','-','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembayaran Piutang','0','" & TextBox3.Text & "','" & hitung & "','LUNAS')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                Dim ubahstatus As String = "update tb_bukupiutang set Status='LUNAS' where ID='" & Label1.Text & "'"
                CMD = New OdbcCommand(ubahstatus, CONN)
                CMD.ExecuteNonQuery()
                'simpan yang baru BELUM
                DataGridView2.DataSource = Nothing
                tampilisi2()
                Dim ss As Double = CDbl(Label10.Text) - CDbl(TextBox4.Text)
                Dim sld As Double = CDbl(DataGridView2.Item(7, 0).Value) + ss
                Dim simpanlagi As String = "insert into tb_bukupiutang values ('','-','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembayaran Piutang','" & ss & "','0','" & sld & "','BELUM LUNAS')"
                CMD = New OdbcCommand(simpanlagi, CONN)
                CMD.ExecuteNonQuery()

                DataGridView2.DataSource = Nothing
                Call tampilisikas()
                Dim hitungsaldo As Double
                hitungsaldo = CDbl(DataGridView2.Item(6, 0).Value) + CDbl(TextBox4.Text)
                Dim simpankas As String = "insert into tb_bukukas values ('','" & TextBox5.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Pembayaran Piutang','" & TextBox4.Text & "','0','" & hitungsaldo & "')"
                CMD = New OdbcCommand(simpankas, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Pembayaran Berhasil", vbInformation + vbOKOnly, "Informasi")
                Call bersih()
                Call mati()
                Button1.Enabled = True
                DataGridView1.DataSource = Nothing
                Call tampilisi()
            End If
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged
        If Not IsNumeric(TextBox6.Text) Then
            TextBox6.Text = 0
        End If
        If TextBox3.Text = "" Then
            TextBox3.Text = 0
        End If
        Label10.Text = CDbl(TextBox3.Text) * CDbl(TextBox6.Text) / 100
        Label10.Text = CDbl(TextBox3.Text) - CDbl(Label10.Text)
    End Sub
End Class