Public Class InputKeywordForm
    Dim Filter As FilterClass
    Dim Mask() As Char = {"~", "(", ")", "#", "\", _
                          "/", "=", ">", "<", "+", _
                          "-", "*", "%", "&", "|", _
                          "^", "'", """", "[", "]"}
    Dim Keyword As String
    Sub New(ByRef Keyword As String)
        ' 此為 Windows Form 設計工具所需的呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。
        Filter = FilterModule.Filter(Keyword)
        NameLabel.Text = Keyword
        Me.Keyword = Keyword
    End Sub

    Private Sub InputKeywordForm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Main.HeadTabControl.Enabled = True
        Main.TryUpdateButton.Enabled = True
    End Sub

    Private Sub InputKeywordForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            KeywordListBox.Items.AddRange(Filter.Keyword)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub AddKeywordButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddKeywordButton.Click
        Dim Str As String = KeywordFilterTextBox.Text.Trim(" ")
        Dim illegalMask As String = String.Empty
        For Each i As Char In Mask
            If Str.IndexOf(i) >= 0 Then
                illegalMask &= i & " "
            End If
        Next
        If illegalMask.Length > 0 Then
            MsgBox("抱歉，您輸入的值含有非法字元：" & illegalMask, MsgBoxStyle.Critical, "輸入錯誤")
            Exit Sub
        End If
        Dim Input() As String = Str.Split(New Char() {vbTab, " "})
        For Each i As String In Input
            If KeywordListBox.FindStringExact(i) = ListBox.NoMatches AndAlso Not String.IsNullOrEmpty(i) Then
                Filter.Keyword = i
                KeywordListBox.Items.Add(i)
                Main.KeywordFilterListBox.Items.Add(i)
            End If
        Next
        KeywordFilterTextBox.Text = String.Empty
    End Sub

    Private Sub DelSelectKeywordButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelSelectKeywordButton.Click
        Del()
    End Sub

    Sub Del()
        If KeywordListBox.SelectedIndex = ListBox.NoMatches Then
            MsgBox("請選擇要刪除的項目")
        Else
            Filter.Remove(FilterEnum.Keyword, KeywordListBox.SelectedItem)
            Main.KeywordFilterListBox.Items.RemoveAt(KeywordListBox.SelectedIndex)
            KeywordListBox.Items.RemoveAt(KeywordListBox.SelectedIndex)
        End If
    End Sub

    Private Sub CloseButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseButton.Click
        If KeywordListBox.Items.Count = 0 Then
            FilterModule.Remove(Keyword)
        End If
        Me.Close()
    End Sub '關閉

    Private Sub KeywordFilterTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeywordFilterTextBox.TextChanged
        AddKeywordButton.Enabled = IIf(String.IsNullOrEmpty(sender.Text.Trim), False, True)
    End Sub

    Private Sub illegalCharButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles illegalCharButton.Click
        Dim MaskStr As String = String.Empty
        For Each i As Char In Mask
            MaskStr &= i & " "
        Next
        MsgBox("查詢時輸入以下字元會導致查詢錯誤" & vbNewLine & _
               "例如：" & MaskStr)
    End Sub

    Private Sub KeywordListBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles KeywordListBox.KeyUp
        If e.KeyCode = Keys.Delete AndAlso KeywordListBox.SelectedItem IsNot Nothing Then
            Del()
        End If
    End Sub
End Class