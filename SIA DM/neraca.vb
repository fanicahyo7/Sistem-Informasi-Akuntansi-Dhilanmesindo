Imports System.Data.Odbc
Public Class neraca
    Sub keluar()
        Me.Close()
    End Sub
    Private Sub neraca_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        DateTimePicker1.Value = Now
    End Sub
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From tb_neraca where tanggal ='" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'", CONN)
        DS = New DataSet
        DA.Fill(DS, "neraca")
        DataGridView1.DataSource = DS.Tables("neraca")
        DataGridView1.ReadOnly = True
    End Sub
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim aktiva, pasiva, saldo As Double
        Call Koneksi()

        Dim tanggal As String = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        Dim bulan, tahun As String
        bulan = DateTimePicker1.Value.Month
        tahun = DateTimePicker1.Value.Year

        DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal

        Call TampilGrid()
        'Dim tanggal, bulan, tahun As String
        'bulan = DateTimePicker1.Value.Month
        'tahun = DateTimePicker1.Value.Year
        'tanggal = 1
        'DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal
        CMD = New OdbcCommand("select sum(Saldo_Aktiva), sum(Saldo_Pasiva) FROM tb_neraca where tanggal ='" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()

        If RD.HasRows Then
            If IsDBNull(RD.Item(0)) Then
                aktiva = 0
            Else
                aktiva = RD.Item(0)
            End If

            If IsDBNull(RD.Item(1)) Then
                pasiva = 0
            Else
                pasiva = RD.Item(1)
            End If
        End If
        saldo = aktiva - pasiva
        Label1.Text = aktiva
        Label2.Text = pasiva
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim tanggal As String = Hasil(Year(DateTimePicker1.Text), Month(DateTimePicker1.Text))
        Dim bulan, tahun As String
        bulan = DateTimePicker1.Value.Month
        tahun = DateTimePicker1.Value.Year

        DateTimePicker2.Value = tahun + "-" + bulan + "-" + tanggal
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        reviewkas.CrystalReportViewer1.RefreshReport()
        reviewneraca.CrystalReportViewer1.SelectionFormula = "ToText({tb_neraca1.Tanggal})='" & CDate(Format(DateTimePicker2.Value, "dd-MM-yyyy")) & "'"
        reviewneraca.CrystalReportViewer1.Refresh()
        reviewneraca.Show()
    End Sub
End Class