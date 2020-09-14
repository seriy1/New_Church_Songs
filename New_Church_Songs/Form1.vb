Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetRegistry()
        Dim loaddef As Boolean
        loaddef = m_LoadDefaults()
        If loaddef Then
            LocateOfficeDir()
        End If
        LocateOfficeDir()
        SaveRegistry()
    End Sub
End Class
