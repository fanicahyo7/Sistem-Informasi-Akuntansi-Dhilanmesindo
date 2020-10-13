Imports System.Data.Odbc
Public Class frmbukupembelian
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From tb_bukupembelian", CONN)
        DS = New DataSet
        DA.Fill(DS, "beli")
        DataGridView1.DataSource = DS.Tables("beli")
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub frmbukupembelian_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call TampilGrid()
    End Sub
End Class