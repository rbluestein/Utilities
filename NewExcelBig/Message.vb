Public Class Message
    Private Sub Message_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtMessage.Text = "Success"
        Timer1.Interval = 1000
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        Me.Close()
    End Sub
End Class