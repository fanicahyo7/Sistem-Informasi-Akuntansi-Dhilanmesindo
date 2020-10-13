Imports System.Data.Odbc
Public Class frmmodal
    Function Hasil(ByVal MyYear As Integer, ByVal MyMonth As Integer) As Integer
        Return DateTime.DaysInMonth(MyYear, MyMonth)
    End Function
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        DateTimePicker1.Value = Now
        If TextBox1.Text = "" Then
            MsgBox("Modal Belum Diisi!", vbCritical + vbOKOnly, "Peringatan")
        Else
            Dim simpansaldomodal As String = "insert into tb_modal values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Modal Awal','" & TextBox1.Text & "')"
            CMD = New OdbcCommand(simpansaldomodal, CONN)
            CMD.ExecuteNonQuery()
            Dim simpankas As String = "insert into tb_bukukas values ('','MD0001','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Modal Awal','" & TextBox1.Text & "','0','" & TextBox1.Text & "')"
            CMD = New OdbcCommand(simpankas, CONN)
            CMD.ExecuteNonQuery()
            MsgBox("Modal Berhasil Tersimpan", vbInformation + vbOKOnly, "Informasi")
            Me.Dispose()
        End If
    End Sub

    Private Sub frmmodal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2
        DateTimePicker1.Value = Now
    End Sub

    Private Sub Button1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Button1.KeyPress
        DateTimePicker1.Value = Now
        If (e.KeyChar = Chr(13)) Then
            If TextBox1.Text = "" Then
                MsgBox("Modal Belum Diisi!", vbCritical + vbOKOnly, "Peringatan")
            Else
                Dim simpansaldomodal As String = "insert into tb_modal values ('','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Modal Awal','" & TextBox1.Text & "')"
                CMD = New OdbcCommand(simpansaldomodal, CONN)
                CMD.ExecuteNonQuery()
                Dim simpankas As String = "insert into tb_bukukas values ('','MD0001','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','Modal Awal','" & TextBox1.Text & "','0','" & TextBox1.Text & "')"
                CMD = New OdbcCommand(simpankas, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Modal Berhasil Tersimpan", vbInformation + vbOKOnly, "Informasi")
                Me.Dispose()
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class