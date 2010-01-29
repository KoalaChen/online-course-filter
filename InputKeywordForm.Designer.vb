<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InputKeywordForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InputKeywordForm))
        Me.KeywordFilterTextBox = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.KeywordListBox = New System.Windows.Forms.ListBox
        Me.AddKeywordButton = New System.Windows.Forms.Button
        Me.DelSelectKeywordButton = New System.Windows.Forms.Button
        Me.NameLabel = New System.Windows.Forms.Label
        Me.CloseButton = New System.Windows.Forms.Button
        Me.illegalCharButton = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'KeywordFilterTextBox
        '
        Me.KeywordFilterTextBox.Location = New System.Drawing.Point(14, 12)
        Me.KeywordFilterTextBox.Name = "KeywordFilterTextBox"
        Me.KeywordFilterTextBox.Size = New System.Drawing.Size(210, 22)
        Me.KeywordFilterTextBox.TabIndex = 0
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(12, 37)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(173, 12)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "．輸入多筆關鍵字請用空白區隔"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(171, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 12)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "目前已鍵入關鍵字"
        '
        'KeywordListBox
        '
        Me.KeywordListBox.FormattingEnabled = True
        Me.KeywordListBox.ItemHeight = 12
        Me.KeywordListBox.Location = New System.Drawing.Point(14, 89)
        Me.KeywordListBox.Name = "KeywordListBox"
        Me.KeywordListBox.Size = New System.Drawing.Size(258, 124)
        Me.KeywordListBox.TabIndex = 7
        '
        'AddKeywordButton
        '
        Me.AddKeywordButton.Enabled = False
        Me.AddKeywordButton.Location = New System.Drawing.Point(230, 10)
        Me.AddKeywordButton.Name = "AddKeywordButton"
        Me.AddKeywordButton.Size = New System.Drawing.Size(42, 23)
        Me.AddKeywordButton.TabIndex = 1
        Me.AddKeywordButton.Text = "新增"
        Me.AddKeywordButton.UseVisualStyleBackColor = True
        '
        'DelSelectKeywordButton
        '
        Me.DelSelectKeywordButton.Location = New System.Drawing.Point(182, 219)
        Me.DelSelectKeywordButton.Name = "DelSelectKeywordButton"
        Me.DelSelectKeywordButton.Size = New System.Drawing.Size(42, 23)
        Me.DelSelectKeywordButton.TabIndex = 8
        Me.DelSelectKeywordButton.Text = "刪除"
        Me.DelSelectKeywordButton.UseVisualStyleBackColor = True
        '
        'NameLabel
        '
        Me.NameLabel.AutoSize = True
        Me.NameLabel.Location = New System.Drawing.Point(12, 74)
        Me.NameLabel.Name = "NameLabel"
        Me.NameLabel.Size = New System.Drawing.Size(44, 12)
        Me.NameLabel.TabIndex = 5
        Me.NameLabel.Text = "<Name>"
        '
        'CloseButton
        '
        Me.CloseButton.Location = New System.Drawing.Point(230, 219)
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Size = New System.Drawing.Size(42, 23)
        Me.CloseButton.TabIndex = 9
        Me.CloseButton.Text = "關閉"
        Me.CloseButton.UseVisualStyleBackColor = True
        '
        'illegalCharButton
        '
        Me.illegalCharButton.Location = New System.Drawing.Point(179, 48)
        Me.illegalCharButton.Name = "illegalCharButton"
        Me.illegalCharButton.Size = New System.Drawing.Size(44, 23)
        Me.illegalCharButton.TabIndex = 4
        Me.illegalCharButton.Text = "說明"
        Me.illegalCharButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(161, 12)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "．輸入的值不能包含以下字元"
        '
        'InputKeywordForm
        '
        Me.AcceptButton = Me.AddKeywordButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 250)
        Me.Controls.Add(Me.illegalCharButton)
        Me.Controls.Add(Me.CloseButton)
        Me.Controls.Add(Me.NameLabel)
        Me.Controls.Add(Me.DelSelectKeywordButton)
        Me.Controls.Add(Me.AddKeywordButton)
        Me.Controls.Add(Me.KeywordListBox)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.KeywordFilterTextBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "InputKeywordForm"
        Me.Text = "關鍵字"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents KeywordFilterTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents KeywordListBox As System.Windows.Forms.ListBox
    Friend WithEvents AddKeywordButton As System.Windows.Forms.Button
    Friend WithEvents DelSelectKeywordButton As System.Windows.Forms.Button
    Friend WithEvents NameLabel As System.Windows.Forms.Label
    Friend WithEvents CloseButton As System.Windows.Forms.Button
    Friend WithEvents illegalCharButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
