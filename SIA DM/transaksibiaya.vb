Imports System.Data.Odbc
Public Class transaksibiaya
    Sub kodeotomatis()
        CMD = New OdbcCommand("select * from tb_bukubiaya order by Kode_Transaksi desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox3.Text = "TY" + "0001"
        Else
            TextBox3.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Transaksi").ToString, 4, 3)) + 1
            If Len(TextBox3.Text) = 1 Then
                TextBox3.Text = "TY000" & TextBox3.Text & ""
            ElseIf Len(TextBox1.Text) = 2 Then
                TextBox3.Text = "TY00" & TextBox3.Text & ""
            ElseIf Len(TextBox1.Text) = 3 Then
                TextBox3.Text = "TY0" & TextBox3.Text & ""
            End If
        End If
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        DateTimePicker1.Value = Now
    End Sub
    Sub isimati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        DateTimePicker1.Enabled = False
    End Sub
    Sub isihidup()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        DateTimePicker1.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call bersih()
        Call kodeotomatis()
        Call isihidup()
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
    End Sub
    Sub keluar()
        Me.Close()
    End Sub
    Private Sub transaksibiaya_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Call Koneksi()
        Call isimati()
        Button2.Enabled = False
        Button3.Enabled = False
        TampilGridKas()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Call isimati()
        Call bersih()
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
    End Sub
    Sub TampilGridKas()
        DA = New OdbcDataAdapter("select * From tb_bukubiaya order by kode_transaksi desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "biaya")
        DataGridView1.DataSource = DS.Tables("biaya")
        DataGridView1.ReadOnly = True
    End Sub
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim tanggal As Integer
        tanggal = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan")
        Else
            If DateTimePicker1.Value.Day = tanggal Then
                Dim simpan As String = "insert into tb_bukubiaya values ('" & TextBox3.Text & "','" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "','" & TextBox1.Text & "','" & TextBox2.Text & "')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            Else
                Dim simpan As String = "insert into tb_bukubiaya values ('" & TextBox3.Text & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & TextBox1.Text & "','" & TextBox2.Text & "')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
            End If
            MsgBox("Data Berhasil Disimpan", vbInformation + vbOKOnly, "Informasi")
            Call bersih()
            Call isimati()
            Button1.Enabled = True
            Button2.Enabled = False
            Button3.Enabled = False
            DataGridView1.DataSource = Nothing
            TampilGridKas()
        End If
    End Sub
End Class