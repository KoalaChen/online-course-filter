Public Class ShowLoading

    Private Sub ShowLoading_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        LoadingTimer.Stop()
    End Sub

    Private Sub ShowLoading_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RemainTimeLabel.Text = 30
        LoadingTimer.Start()
    End Sub

    Private Sub LoadingTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadingTimer.Tick
        Dim Time As Short = RemainTimeLabel.Text
        If Time - 1 >= 0 Then
            Time -= 1
            LoadingProgressBar.Value = (30 - Time) / 30 * 100
            RemainTimeLabel.Text = Time
        Else
            TimeOut = True
            Me.Close()
        End If
    End Sub
End Class