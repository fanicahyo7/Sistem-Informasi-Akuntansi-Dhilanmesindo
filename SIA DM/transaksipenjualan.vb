Imports System.Data.Odbc
Public Class transaksipenjualan
    Public baris As Integer = 0
    Sub kodeotomatis()
        CMD = New OdbcCommand("select * from tb_bukupenjualan order by Kode_Transaksi desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox1.Text = "TJ" + "0001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("Kode_Transaksi").ToString, 4, 3)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "TJ000" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 2 Then
                TextBox1.Text = "TJ00" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 3 Then
                TextBox1.Text = "TJ0" & TextBox1.Text & ""
            End If
        End If
    End Sub
    Sub isicombo()
        CMD = New OdbcCommand("select Kode_Barang FROM tb_barang", CONN)
        RD = CMD.ExecuteReader
        Do While RD.Read
            ComboBox1.Items.Add(RD.Item(0))
        Loop
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        With ListView1
            .Items.Add(baris + 1)
            .Items(baris).SubItems.Add(ComboBox1.Text)
            .Items(baris).SubItems.Add(TextBox2.Text)
            .Items(baris).SubItems.Add(TextBox3.Text)
            .Items(baris).SubItems.Add(TextBox4.Text)
            .Items(baris).SubItems.Add(TextBox5.Text)
            .Items(baris).SubItems.Add(TextBox6.Text)
        End With
        Label10.Text += CInt(TextBox6.Text)
        baris += 1
        ComboBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = 0
        TextBox5.Text = 0
        TextBox6.Text = 0
    End Sub
    Sub TampilGridKas()
        DA = New OdbcDataAdapter("select * From tb_bukukas order by id desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "kas")
        DataGridView1.DataSource = DS.Tables("kas")
        DataGridView1.ReadOnly = True
    End Sub
    Sub TampilGridpenjualan()
        DA = New OdbcDataAdapter("select * From tb_bukupenjualan order by Kode_Transaksi desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "jual")
        DataGridView1.DataSource = DS.Tables("jual")
        DataGridView1.ReadOnly = True
    End Sub
    Sub TampilGridpiutang()
        DA = New OdbcDataAdapter("select * From tb_bukupiutang order by ID desc", CONN)
        DS = New DataSet
        DA.Fill(DS, "piutang")
        DataGridView1.DataSource = DS.Tables("piutang")
        DataGridView1.ReadOnly = True
    End Sub
    Sub isimati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        ComboBox1.Enabled = False
        DateTimePicker1.Enabled = False
        Button1.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
    End Sub
    Sub isihidup()
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        DateTimePicker1.Enabled = True
        ComboBox1.Enabled = True
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        DateTimePicker1.Value = Now
        ComboBox1.Text = ""
        Label10.Text = 0
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Call bersih()
        Call kodeotomatis()
        Call isihidup()
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = True
        Button4.Enabled = True
        TextBox4.Text = 0
        TextBox5.Text = 0
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
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Call isimati()
        Call bersih()
        Button2.Enabled = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        CMD = New OdbcCommand("select * FROM tb_barang where Kode_Barang= '" & ComboBox1.Text & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        TextBox2.Text = RD.Item(1)
        TextBox3.Text = RD.Item(2)
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged
        If Not IsNumeric(TextBox4.Text) Then
            TextBox4.Text = 0
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged
        If Not IsNumeric(TextBox5.Text) Then
            TextBox5.Text = 0
        End If
        TextBox6.Text = TextBox5.Text * TextBox4.Text
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        bayarpenjualan.Show()
        bayarpenjualan.Label3.Text = Label10.Text
    End Sub
    Sub keluar()
        Me.Close()
    End Sub
    Private Sub transaksipenjualan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Call Koneksi()
        Call isimati()
        Call bersih()
        Call isicombo()
        Call kolomlist()
    End Sub
End Class