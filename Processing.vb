Public Class Processing
    Private Sub Processing_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Counting()
    End Sub
    Sub Counting()
        If StartPage = 0 AndAlso EndPage = 0 Then
            Label2.Text = "等待中..."
            ProgressBar1.Value = 0
        ElseIf Main.HeadTabControl.Enabled OrElse TimeOut Then
            ProgressBar1.Value = 100
            Me.Close()
        Else
            ProgressBar1.Value = IIf(EndPage <> 0, (StartPage / EndPage) * 100, 0)
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Counting()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sender.Text = "停止中"
        StopProcessing = True
    End Sub
End Class