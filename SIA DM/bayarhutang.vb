Imports System.Data.Odbc
Public Class bayarhutang
    Sub tampilisi()
        DA = New OdbcDataAdapter("select Kode_Hutang,Tanggal,Keterangan,Kredit as Hutang from tb_bukuhutang where Status='BELUM LUNAS'", CONN)
        DS = New DataSet
        DA.Fill(DS, "utang")
        DataGridView1.DataSource = DS.Tables("utang")
        DataGridView1.ReadOnly = True
        DataGridView1.Refresh()
    End Sub
    Sub tampilisi2()
        DA = New OdbcDataAdapter("select * from tb_bukuhutang order by kode_hutang desc", CONN)
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
    Sub kodeotomatis()
        CMD = New OdbcCommand("select * from tb_bukuhutang order by kode_hutang desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox5.Text = "BH" + "0001"
        Else
            TextBox5.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Hutang").ToString, 4, 3)) + 1
            If Len(TextBox5.Text) = 1 Then
                TextBox5.Text = "BH000" & TextBox5.Text & ""
            ElseIf Len(TextBox5.Text) = 2 Then
                TextBox5.Text = "BH00" & TextBox5.Text & ""
            ElseIf Len(TextBox5.Text) = 3 Then
                TextBox5.Text = "BH0" & TextBox5.Text & ""
            End If
        End If
    End Sub
    Private Sub bayarhutang_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Call Koneksi()
        Call tampilisi()
        Call mati()
    End Sub
    Sub keluar()
        Me.Close()
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        DateTimePicker1.Value = Now
        DateTimePicker2.Value = Now
    End Sub
    Sub mati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        DateTimePicker1.Enabled = False
        DateTimePicker2.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        For f As Integer = 0 To DataGridView1.RowCount - 1
            If IsDBNull(DataGridView1.CurrentRow.Cells(0).Value) Then
                TextBox1.Text = ""
            Else
                TextBox1.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(1).Value) Then
                DateTimePicker1.Value = Now
            Else
                DateTimePicker1.Value = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(2).Value) Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(3).Value) Then
                TextBox3.Text = ""
            Else
                TextBox3.Text = DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih Data yang Akan Dibayar", vbCritical + vbOKOnly, "Peringatan")
        Else
            TextBox4.Enabled = True
            Button1.Enabled = False
            Button2.Enabled = True
            Button3.Enabled = True
            Call kodeotomatis()
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
        tanggal = Hasil(Year(DateTimePicker2.Text), Month(DateTimePicker2.Text))
        If DateTimePicker2.Value.Day = tanggal Then
            DataGridView2.DataSource = Nothing
            DateTimePicker1.Value = Now
            Call tampilisi2()
            Dim hitung As Double
            hitung = CDbl(DataGridView2.Item(5, 0).Value) - CDbl(TextBox4.Text)
            Dim simpan As String = "insert into tb_bukuhutang values ('" & TextBox5.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker2.Value), "yyyy-MM-dd") & "','Pembayaran " & TextBox1.Text & "','" & TextBox4.Text & "','0','" & hitung & "','LUNAS')"
            CMD = New OdbcCommand(simpan, CONN)
            CMD.ExecuteNonQuery()
            Dim ubahstatus As String = "update tb_bukuhutang set Status='LUNAS' where kode_hutang='" & TextBox1.Text & "'"
            CMD = New OdbcCommand(ubahstatus, CONN)
            CMD.ExecuteNonQuery()

            DataGridView2.DataSource = Nothing
            Call tampilisikas()
            Dim hitungsaldo As Double
            hitungsaldo = CDbl(DataGridView2.Item(6, 0).Value) - CDbl(TextBox4.Text)
            Dim simpankas As String = "insert into tb_bukukas values ('','" & TextBox5.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker2.Value), "yyyy-MM-dd") & "','Pembayaran Hutang','0','" & TextBox4.Text & "','" & hitungsaldo & "')"
            CMD = New OdbcCommand(simpankas, CONN)
            CMD.ExecuteNonQuery()
            MsgBox("Pembayaran Berhasil", vbInformation + vbOKOnly, "Informasi")
            Call bersih()
            Call mati()
            Button1.Enabled = True
            DataGridView1.DataSource = Nothing
            Call tampilisi()
        Else
            DataGridView2.DataSource = Nothing
            DateTimePicker1.Value = Now
            Call tampilisi2()
            Dim hitung As Double
            hitung = CDbl(DataGridView2.Item(5, 0).Value) - CDbl(TextBox4.Text)
            Dim simpan As String = "insert into tb_bukuhutang values ('" & TextBox5.Text & "','" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "','Pembayaran " & TextBox1.Text & "','" & TextBox4.Text & "','0','" & hitung & "','LUNAS')"
            CMD = New OdbcCommand(simpan, CONN)
            CMD.ExecuteNonQuery()
            Dim ubahstatus As String = "update tb_bukuhutang set Status='LUNAS' where kode_hutang='" & TextBox1.Text & "'"
            CMD = New OdbcCommand(ubahstatus, CONN)
            CMD.ExecuteNonQuery()

            DataGridView2.DataSource = Nothing
            Call tampilisikas()
            Dim hitungsaldo As Double
            hitungsaldo = CDbl(DataGridView2.Item(6, 0).Value) - CDbl(TextBox4.Text)
            Dim simpankas As String = "insert into tb_bukukas values ('','" & TextBox5.Text & "','" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "','Pembayaran Hutang','0','" & TextBox4.Text & "','" & hitungsaldo & "')"
            CMD = New OdbcCommand(simpankas, CONN)
            CMD.ExecuteNonQuery()
            MsgBox("Pembayaran Berhasil", vbInformation + vbOKOnly, "Informasi")
            Call bersih()
            Call mati()
            Button1.Enabled = True
            DataGridView1.DataSource = Nothing
            Call tampilisi()
        End If
    End Sub
End Class