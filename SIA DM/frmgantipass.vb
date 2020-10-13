Imports System.Data.Odbc
Public Class frmgantipass

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dispose()
    End Sub

    Private Sub frmgantipass_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Form1
        keTengah(Form1, Me)
        Koneksi()
        TextBox2.Enabled = False
        TextBox3.Enabled = False
    End Sub
    Sub keluar()
        Me.Close()
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        CMD = New OdbcCommand("select * from tb_user where Username='admin'", CONN)
        RD = CMD.ExecuteReader()
        RD.Read()
        Dim pass As String
        pass = RD.Item(1)

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Data Belum Lengkap", vbCritical + vbOKOnly, "Peringatan")
        ElseIf TextBox1.Text = pass Then
            MsgBox("Password Lama Salah!", vbCritical + vbOKOnly, "Peringatan")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
        ElseIf Not TextBox2.Text = TextBox3.Text Then
            MsgBox("Password Baru Salah!", vbCritical + vbOKOnly, "Peringatan")
            TextBox2.Text = ""
            TextBox3.Text = ""
        Else
            Dim ubahpass As String = "update tb_user set Password='" & TextBox3.Text & "' where Username='admin'"
            CMD = New OdbcCommand(ubahpass, CONN)
            CMD.ExecuteNonQuery()
            MsgBox("Password Berhasil Diubah", vbInformation + vbOKOnly, "Informasi")
            Me.Close()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If (e.KeyChar = Chr(13)) Then
            CMD = New OdbcCommand("select * from tb_user where Username='admin'", CONN)
            RD = CMD.ExecuteReader()
            RD.Read()
            Dim pass As String
            pass = RD.Item(1)
            If TextBox1.Text = pass Then
                TextBox2.Enabled = True
                TextBox3.Enabled = True
                TextBox2.Focus()
            Else
                MsgBox("Password Lama Salah!", vbCritical + vbOKOnly, "Peringatan")
                TextBox1.Text = ""
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If (e.KeyChar = Chr(13)) Then
            If TextBox2.Text = TextBox3.Text Then
                Dim ubahpass As String = "update tb_user set Password='" & TextBox3.Text & "' where Username='admin'"
                CMD = New OdbcCommand(ubahpass, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Password Berhasil Diubah", vbInformation + vbOKOnly, "Informasi")
                Me.Close()
            Else
                MsgBox("Password Baru Salah!", vbCritical + vbOKOnly, "Peringatan")
                TextBox2.Text = ""
                TextBox3.Text = ""
            End If
        End If
    End Sub
End Class