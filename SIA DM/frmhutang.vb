Imports System.Data.Odbc
Public Class frmhutang
    Sub kodeotomatis()
        CMD = New OdbcCommand("select * from tb_bukuhutang order by kode_hutang desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox1.Text = "BH" + "0001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Hutang").ToString, 4, 3)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "BH000" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 2 Then
                TextBox1.Text = "BH00" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 3 Then
                TextBox1.Text = "BH0" & TextBox1.Text & ""
            End If
        End If
    End Sub
    Private Sub frmhutang_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Koneksi()
        bersih()
        mati()
    End Sub
    Sub keluar()
        Me.Close()
    End Sub
    Sub mati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        DateTimePicker1.Enabled = False
        Button1.Enabled = True
    End Sub
    Sub hidup()
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        DateTimePicker1.Enabled = True
        Button1.Enabled = False
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        DateTimePicker1.Value = Now
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        bersih()
        kodeotomatis()
        hidup()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        bersih()
        mati()
    End Sub
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim tanggal As Integer
        tanggal = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        If DateTimePicker1.Value.Day = tanggal Then
            DataGridView2.DataSource = Nothing
            Call TampilGridhutang()
            If DataGridView2.RowCount = 0 Then
                Dim simpan7 As String = "insert into tb_bukuhutang values ('" & TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','" & TextBox2.Text & "','0','" & TextBox3.Text & "','" & TextBox3.Text & "','BELUM LUNAS')"
                CMD = New OdbcCommand(simpan7, CONN)
                CMD.ExecuteNonQuery()
            Else
                Dim hitung As Double
                hitung = CDbl(DataGridView2.Item(5, 0).Value) + CDbl(TextBox3.Text)
                Dim simpan7 As String = "insert into tb_bukuhutang values ('" & TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','" & TextBox2.Text & "','0','" & TextBox3.Text & "','" & hitung & "','BELUM LUNAS')"
                CMD = New OdbcCommand(simpan7, CONN)
                CMD.ExecuteNonQuery()
            End If

            DataGridView2.DataSource = Nothing
            Call TampilGridkas()
            If DataGridView2.RowCount = 0 Then
                Dim simpan As String = "insert into tb_bukukas values ('','" & TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','" & TextBox2.Text & "','" & TextBox3.Text & "','0','" & TextBox3.Text & "')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Else
                Dim hitung As Double
                hitung = CDbl(DataGridView2.Item(6, 0).Value) + CDbl(TextBox3.Text)
                Dim simpan As String = "insert into tb_bukukas values ('','" & TextBox1.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','" & TextBox2.Text & "','" & TextBox3.Text & "','0','" & hitung & "')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            End If
            MsgBox("Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi")
            bersih()
            mati()
        Else
            DataGridView2.DataSource = Nothing
            Call TampilGridhutang()
            If DataGridView2.RowCount = 0 Then
                Dim simpan7 As String = "insert into tb_bukuhutang values ('" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','0','" & TextBox3.Text & "','" & TextBox3.Text & "','BELUM LUNAS')"
                CMD = New OdbcCommand(simpan7, CONN)
                CMD.ExecuteNonQuery()
            Else
                Dim hitung As Double
                hitung = CDbl(DataGridView2.Item(5, 0).Value) + CDbl(TextBox3.Text)
                Dim simpan7 As String = "insert into tb_bukuhutang values ('" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','0','" & TextBox3.Text & "','" & hitung & "','BELUM LUNAS')"
                CMD = New OdbcCommand(simpan7, CONN)
                CMD.ExecuteNonQuery()
            End If

            DataGridView2.DataSource = Nothing
            Call TampilGridkas()
            If DataGridView2.RowCount = 0 Then
                Dim simpan As String = "insert into tb_bukukas values ('','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','" & TextBox3.Text & "','0','" & TextBox3.Text & "')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Else
                Dim hitung As Double
                hitung = CDbl(DataGridView2.Item(6, 0).Value) + CDbl(TextBox3.Text)
                Dim simpan As String = "insert into tb_bukukas values ('','" & TextBox1.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & TextBox2.Text & "','" & TextBox3.Text & "','0','" & hitung & "')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            End If
            MsgBox("Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi")
            bersih()
            mati()
        End If
    End Sub
    Sub TampilGridhutang()
        DA = New OdbcDataAdapter("select * From tb_bukuhutang order by kode_hutang desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "hutang")
        DataGridView2.DataSource = DS.Tables("hutang")
        DataGridView2.ReadOnly = True
    End Sub
    Sub TampilGridkas()
        DA = New OdbcDataAdapter("select * From tb_bukukas order by id desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "kas")
        DataGridView2.DataSource = DS.Tables("kas")
        DataGridView2.ReadOnly = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class