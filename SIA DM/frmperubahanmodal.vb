Imports System.Data.Odbc
Public Class frmperubahanmodal
    Private Sub frmperubahanmodal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        DateTimePicker1.Value = Now
    End Sub
    Sub keluar()
        Me.Close()
    End Sub
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From tb_modal where tanggal ='" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'", CONN)
        DS = New DataSet
        DA.Fill(DS, "labarugi")
        DataGridView1.DataSource = DS.Tables("labarugi")
        DataGridView1.ReadOnly = True
    End Sub
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function

    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim tanggal As String = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        Dim bulan, tahun As String
        bulan = DateTimePicker1.Value.Month
        tahun = DateTimePicker1.Value.Year

        DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call Koneksi()

        Dim tanggal As String = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        Dim bulan, tahun As String
        bulan = DateTimePicker1.Value.Month
        tahun = DateTimePicker1.Value.Year

        DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal

        Call TampilGrid()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        reviewkas.CrystalReportViewer1.RefreshReport()
        reviewperubahanmodal.CrystalReportViewer1.SelectionFormula = "ToText({tb_modal1.Tanggal})='" & CDate(Format(DateTimePicker2.Value, "dd-MM-yyyy")) & "'"
        reviewperubahanmodal.CrystalReportViewer1.Refresh()
        reviewperubahanmodal.Show()
    End Sub
End Class