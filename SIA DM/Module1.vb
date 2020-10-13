Imports System.Data.Odbc
Module Module1
    Public CONN As OdbcConnection
    Public CMD As OdbcCommand
    Public DS As New DataSet
    Public DA As OdbcDataAdapter
    Public RD As OdbcDataReader
    Dim LokasiData As String
    Sub Koneksi()
        LokasiData = "Driver={MySQL ODBC 3.51 Driver};Database=db_sia;server=localhost;uid=root"
        CONN = New OdbcConnection(LokasiData)
        If CONN.State = ConnectionState.Closed Then
            CONN.Open()
        End If
    End Sub
    Sub keTengah(ByVal parent As Form, ByVal child As Form)
        Dim frmA As Form = parent
        Dim frmB As Form = child
        Dim x As Integer = (parent.Width / 2) - (child.Width / 2)
        Dim y As Integer = (parent.Height / 2.2) - (child.Height / 2)
        child.Location = New Point(x, y)
    End Sub
End Module
