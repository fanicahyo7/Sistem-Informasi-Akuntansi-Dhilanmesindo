Imports System.Data.Odbc
Public Class riwayatpersediaanbarang
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From tb_persediaanbarang", CONN)
        DS = New DataSet
        DA.Fill(DS, "persediaanbarang")
        DataGridView1.DataSource = DS.Tables("persediaanbarang")
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub riwayatpersediaanbarang_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call TampilGrid()
    End Sub
End Class