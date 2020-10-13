Imports System.Data.Odbc
Public Class login

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Masukkan Username Dan Password", vbCritical + vbOKOnly, "Peringatan")
        Else
            CMD = New OdbcCommand("select * from tb_user where Username='" & TextBox1.Text & "' AND Password='" & TextBox2.Text & "'", CONN)
            RD = CMD.ExecuteReader()
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Username dan Password Salah", vbCritical + vbOKOnly, "Peringatan")
                TextBox1.Text = ""
                TextBox2.Text = ""
            Else
                Me.Hide()
                Form1.Show()
            End If
        End If
    End Sub

    Private Sub login_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2
        Call Koneksi()
    End Sub
    Sub keluar()
        Close()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dispose()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If (e.KeyChar = Chr(13)) Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Then
                MsgBox("Masukkan Username Dan Password", vbCritical + vbOKOnly, "Peringatan")
            Else
                CMD = New OdbcCommand("select * from tb_user where Username='" & TextBox1.Text & "' AND Password='" & TextBox2.Text & "'", CONN)
                RD = CMD.ExecuteReader()
                RD.Read()
                If Not RD.HasRows Then
                    MsgBox("Username dan Password Salah", vbCritical + vbOKOnly, "Peringatan")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                Else
                    Me.Hide()
                    Form1.Show()
                End If
            End If
        End If
    End Sub
End Class