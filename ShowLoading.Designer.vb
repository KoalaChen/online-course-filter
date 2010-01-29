<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ShowLoading
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ShowLoading))
        Me.Label1 = New System.Windows.Forms.Label
        Me.RemainTimeLabel = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.LoadingProgressBar = New System.Windows.Forms.ProgressBar
        Me.LoadingTimer = New System.Windows.Forms.Timer(Me.components)
        Me.FirstTimer = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(121, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "網頁擷取中...請稍後"
        '
        'RemainTimeLabel
        '
        Me.RemainTimeLabel.AutoSize = True
        Me.RemainTimeLabel.Location = New System.Drawing.Point(171, 62)
        Me.RemainTimeLabel.Name = "RemainTimeLabel"
        Me.RemainTimeLabel.Size = New System.Drawing.Size(17, 12)
        Me.RemainTimeLabel.TabIndex = 2
        Me.RemainTimeLabel.Text = "00"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(187, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 12)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "秒後停止執行"
        '
        'LoadingProgressBar
        '
        Me.LoadingProgressBar.Location = New System.Drawing.Point(12, 12)
        Me.LoadingProgressBar.Name = "LoadingProgressBar"
        Me.LoadingProgressBar.Size = New System.Drawing.Size(252, 23)
        Me.LoadingProgressBar.TabIndex = 3
        '
        'LoadingTimer
        '
        Me.LoadingTimer.Interval = 1000
        '
        'FirstTimer
        '
        Me.FirstTimer.Enabled = True
        Me.FirstTimer.Interval = 1000
        '
        'ShowLoading
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(276, 90)
        Me.Controls.Add(Me.LoadingProgressBar)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.RemainTimeLabel)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ShowLoading"
        Me.Text = "網頁讀取中"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RemainTimeLabel As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LoadingProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents LoadingTimer As System.Windows.Forms.Timer
    Friend WithEvents FirstTimer As System.Windows.Forms.Timer
End Class
