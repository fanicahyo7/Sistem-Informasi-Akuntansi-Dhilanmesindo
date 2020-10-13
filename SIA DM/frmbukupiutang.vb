Imports System.Data.Odbc
Public Class frmbukupiutang
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From tb_bukupiutang", CONN)
        DS = New DataSet
        DA.Fill(DS, "piutang")
        DataGridView1.DataSource = DS.Tables("piutang")
        DataGridView1.ReadOnly = True
    End Sub
    Sub TampilGridtanggal()
        DA = New OdbcDataAdapter("select * From tb_bukupiutang where tanggal='" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'", CONN)
        DS = New DataSet
        DA.Fill(DS, "piutang")
        DataGridView1.DataSource = DS.Tables("piutang")
        DataGridView1.ReadOnly = True
    End Sub
    Sub TampilGridbulan()
        DA = New OdbcDataAdapter("select * From tb_bukupiutang where tanggal between '" & Format(DateTimePicker3.Value, "yyyy-MM-dd") & "' AND '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'", CONN)
        DS = New DataSet
        DA.Fill(DS, "piutang")
        DataGridView1.DataSource = DS.Tables("piutang")
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub frmbukupiutang_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Call Koneksi()
        RadioButton1.Checked = True
        DateTimePicker1.Value = Now
        TampilGrid()
    End Sub
    Sub keluar()
        Me.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            DataGridView1.DataSource = Nothing
            Call TampilGrid()
        ElseIf RadioButton2.Checked = True Then
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = False
            DataGridView1.DataSource = Nothing
            TampilGridtanggal()
        ElseIf RadioButton3.Checked = True Then
            DateTimePicker2.Enabled = True
            DateTimePicker1.Enabled = False

            Dim tanggal, bulan, tahun As String
            bulan = DateTimePicker2.Value.Month
            tahun = DateTimePicker2.Value.Year
            tanggal = 1
            DateTimePicker3.Value = tahun + "-" + bulan + "-" + tanggal
            Dim km As String = Hasil(Year(DateTimePicker2.Text), Month(DateTimePicker2.Text))
            DateTimePicker2.Value = tahun + "-" + bulan + "-" + km
            TampilGridbulan()
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton1.Checked = True Then
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            DataGridView1.DataSource = Nothing
            Call TampilGrid()
        ElseIf RadioButton2.Checked = True Then
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = False
            DataGridView1.DataSource = Nothing
            TampilGridtanggal()
        ElseIf RadioButton3.Checked = True Then
            DateTimePicker2.Enabled = True
            DateTimePicker1.Enabled = False

            Dim tanggal, bulan, tahun As String
            bulan = DateTimePicker2.Value.Month
            tahun = DateTimePicker2.Value.Year
            tanggal = 1
            DateTimePicker3.Value = tahun + "-" + bulan + "-" + tanggal
            Dim km As String = Hasil(Year(DateTimePicker2.Text), Month(DateTimePicker2.Text))
            DateTimePicker2.Value = tahun + "-" + bulan + "-" + km
            TampilGridbulan()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton1.Checked = True Then
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            DataGridView1.DataSource = Nothing
            Call TampilGrid()
        ElseIf RadioButton2.Checked = True Then
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = False
            DataGridView1.DataSource = Nothing
            TampilGridtanggal()
        ElseIf RadioButton3.Checked = True Then
            DateTimePicker2.Enabled = True
            DateTimePicker1.Enabled = False

            Dim tanggal, bulan, tahun As String
            bulan = DateTimePicker2.Value.Month
            tahun = DateTimePicker2.Value.Year
            tanggal = 1
            DateTimePicker3.Value = tahun + "-" + bulan + "-" + tanggal
            Dim km As String = Hasil(Year(DateTimePicker2.Text), Month(DateTimePicker2.Text))
            DateTimePicker2.Value = tahun + "-" + bulan + "-" + km
            TampilGridbulan()
        End If
    End Sub
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function
    Private Sub DateTimePicker3_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker3.ValueChanged
        TampilGrid()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        TampilGridbulan()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        TampilGridtanggal()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        reviewkas.CrystalReportViewer1.RefreshReport()
        If RadioButton1.Checked = True Then
            reviewpiutang.CrystalReportViewer1.Refresh()
            reviewpiutang.Show()
        ElseIf RadioButton2.Checked = True Then
            reviewpiutang.CrystalReportViewer1.SelectionFormula = "ToText({tb_bukupiutang1.Tanggal})='" & CDate(Format(DateTimePicker1.Value, "dd-MM-yyyy")) & "'"
            reviewpiutang.CrystalReportViewer1.Refresh()
            reviewpiutang.Show()
        ElseIf RadioButton3.Checked = True Then
            reviewpiutang.piutang1.SetParameterValue("awal", CDate(Format(DateTimePicker2.Value, "dd-MM-yyyy")))
            reviewpiutang.piutang1.SetParameterValue("akhir", CDate(Format(DateTimePicker3.Value, "dd-MM-yyyy")))
            reviewpiutang.Show()
        End If
    End Sub
End Class