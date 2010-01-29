'加入參考：Microsoft.mshtml
'已知選擇DropDownList項目->SelectedIndex = 1
'已知選擇Button項目      ->Click()
'Form Post 方式 WebBrowser1.Document.GetElementsByTagName("form").GetElementsByName("form3").Item(0).DomElement.submit()
Imports System.Collections.Specialized
Public Class Main
    Dim Query As System.Windows.Forms.RadioButton '目前Ckecked = True 的 RadioButton 指標
    Dim ComputeValue() As String = {"選取", "時段", "課程大綱", "剩下名額", "課程連結"}
    Dim MultiValue() As String = {"*星期*節次*教室*", "*選上人數*"}

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ShowLoading.Show() '顯示Loading視窗
        Loading(CourseWebBrowser, StatusLabel, CourseUrl) '到課程首頁
        ShowLoading.Close() '關閉Loading視窗
        If CheckPage() Then GetQueryComboBoxItem(1, 8) '檢查錯誤'更新 查詢相關控制項
        '後面篩選
        '星期
        Dim Period() As Char = {"1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8c", "9"c, "A"c, "B"c, "C"c, "D"c}
        For Each i As Char In Period
            WeekAndPeriodDataGridView.Rows.Add(New String() {"第" & i & "節", False, False, False, False, False, False, False})
        Next
        WeekAndPeriodDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
        '人數
        '其他控制項
        For i As Byte = 0 To ResultDataGridView.Columns.Count - 1
            ControlListBox.Items.Add(ResultDataGridView.Columns.Item(i).HeaderText)
        Next
        Dim ControlListBoxExceptItem() As String = {"選取", "時段", "星期", "節次", "人數限制", "選上人數", "課程連結"}
        For Each i As String In ControlListBoxExceptItem
            ControlListBox.Items.Remove(i)
        Next
    End Sub '初始化

    Private Sub RadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    SchoolGeneralEducationRadioButton.CheckedChanged, _
    GeneralEducationRadioButton.CheckedChanged, _
    SportRadioButton.CheckedChanged, _
    MilitaryTrainingRadioButton.CheckedChanged, _
    ResearchInstituteRadioButton.CheckedChanged, _
    AcademicRadioButton.CheckedChanged, _
    EducationRadioButton.CheckedChanged, _
    TimeRadioButton.CheckedChanged, _
    TeachLangRadioButton.CheckedChanged, _
    TeacherNameRadioButton.CheckedChanged, _
    CourseNameRadioButton.CheckedChanged '依類別查詢-RadioButton
        ResearchInstituteComboBox.Enabled = ResearchInstituteRadioButton.Checked
        AcademicPanel.Enabled = AcademicRadioButton.Checked
        TimePanel.Enabled = TimeRadioButton.Checked
        CourseNameTextBox.Enabled = CourseNameRadioButton.Checked
        TeacherNameTextBox.Enabled = TeacherNameRadioButton.Checked
        TeachLangComboBox.Enabled = TeachLangRadioButton.Checked
        Query = sender '記住選到的是哪個RadioButton
    End Sub 'STEP1：查詢

    Sub GetQueryComboBoxItem(ByVal Start As Byte, ByVal [End] As Byte)
        Try
            If Start <= 1 AndAlso 1 <= [End] Then GetComboBoxInformation(ResearchInstituteComboBox, CourseWebBrowser, ComboBoxName("研究所"), 1)
            If Start <= 2 AndAlso 2 <= [End] Then GetComboBoxInformation(AcademicComboBox, CourseWebBrowser, ComboBoxName("學制"), 0)
            If Start <= 3 AndAlso 3 <= [End] Then GetComboBoxInformation(DepartmentComboBox, CourseWebBrowser, ComboBoxName("系別"), 0)
            If Start <= 4 AndAlso 4 <= [End] Then GetComboBoxInformation(GradeComboBox, CourseWebBrowser, ComboBoxName("年級"), 1)
            If Start <= 5 AndAlso 5 <= [End] Then GetComboBoxInformation(ClassComboBox, CourseWebBrowser, ComboBoxName("班別"), 1)
            If Start <= 6 AndAlso 6 <= [End] Then GetComboBoxInformation(WeekComboBox, CourseWebBrowser, ComboBoxName("星期"), 1)
            If Start <= 7 AndAlso 7 <= [End] Then GetComboBoxInformation(PeriodComboBox, CourseWebBrowser, ComboBoxName("節次"), 1)
            If Start <= 8 AndAlso 8 <= [End] Then GetComboBoxInformation(TeachLangComboBox, CourseWebBrowser, ComboBoxName("授課語言"), 1)
        Catch ex As Exception
            FieldUpdate = False
            CheckPage() '顯示錯誤訊息
        End Try
    End Sub '1-8控制項更新

    Sub AcademicComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AcademicComboBox.SelectedIndexChanged
        'ERROR:這裡有發生一個問題，有時系別清單未更新成功
        If AcademicComboBoxIndex = -1 AndAlso sender.SelectedIndex = 0 Then
            AcademicComboBoxIndex = sender.SelectedIndex
            Exit Sub
        End If
        If (Not sender.SelectedIndex = AcademicComboBoxIndex) Then '上次的Index和這次的比較

            '
            '鎖定相關控制項
            '
            AcademicComboBox.Enabled = False
            DepartmentComboBox.Enabled = False
            NextStep2Button.Enabled = False
            ReFreshDataLabel.Visible = True
            '
            '網頁操作
            '
            Try
                Loading(CourseWebBrowser, StatusLabel, CourseUrl) '回首頁
                If PageIsAvailable Then
                    DoComboBoxSelect(CourseWebBrowser, ComboBoxName("學制"), AcademicComboBox.SelectedItem) '選擇
                    CourseWebBrowser.Document.GetElementsByTagName("form").GetElementsByName("form3"). _
                    Item(0).DomElement.submit() '使網頁進行POST動作
                    ReFreshTimer.Start() '更新ComboBox內容
                Else
                    Throw New Exception
                End If
            Catch ex As Exception
                FieldUpdate = False
                CheckPage() '顯示錯誤訊息
            End Try
        End If
        AcademicComboBoxIndex = sender.SelectedIndex '記住這次的Index
    End Sub '學制ComboBox更新

    Private Sub ReFreshTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReFreshTimer.Tick
        sender.Stop() '先停止
        Try
            Dim NewCombobox As New Windows.Forms.ComboBox() '這次的ComboBox應該跟上次不一樣
            GetComboBoxInformation(NewCombobox, CourseWebBrowser, ComboBoxName("系別"), 0)
            If NewCombobox.Items IsNot DepartmentComboBox.Items Then
                DepartmentComboBox.Items.Clear() '清除舊資料
                For Each i As String In NewCombobox.Items
                    DepartmentComboBox.Items.Add(i)
                Next '填入內容
                '
                '解除鎖定相關控制項
                '
                AcademicComboBox.Enabled = True
                DepartmentComboBox.Enabled = True
                NextStep2Button.Enabled = True
                ReFreshDataLabel.Visible = False
            Else
                sender.Start() '若沒更新，啟用
            End If
        Catch ex As Exception
            Throw New Exception '出現錯誤則擲出一個錯誤
        End Try
    End Sub '重抓系別資料

    Private Sub AuthorLinkLabel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles AuthorLinkLabel.LinkClicked
        Try
            System.Diagnostics.Process.Start("http://tomshare.net.ru/")
        Catch ex As Exception
            MsgBox("無法啟動瀏覽器，請在您的瀏覽器輸入網址。")
        End Try
    End Sub '作者的網站

    Private Sub NextStep2Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextStep2Button.Click
        Dim Msg As String = String.Empty '集合錯誤訊息
        HeadTabControl.Enabled = False
        WorkingToolStripProgressBar.Visible = True '顯示進度條
        WorkingToolStripProgressBar.Style = ProgressBarStyle.Marquee
        '檢查資料是否有誤
        Select Case Query.Text
            Case "研究所"
                If ResearchInstituteComboBox.SelectedIndex = -1 Then Msg &= "研究所未選" & vbNewLine
            Case "學制"
                If AcademicComboBox.SelectedIndex = -1 Then Msg &= "學制未選" & vbNewLine
                If DepartmentComboBox.SelectedIndex = -1 Then Msg &= "系別未選" & vbNewLine
                If GradeComboBox.SelectedIndex = -1 Then Msg &= "年級未選" & vbNewLine
                If ClassComboBox.SelectedIndex = -1 Then Msg &= "班別未選" & vbNewLine
            Case "以時間查詢"
                If WeekComboBox.SelectedIndex = -1 Then Msg &= "星期未選" & vbNewLine
                If PeriodComboBox.SelectedIndex = -1 Then Msg &= "節次未選" & vbNewLine
            Case "課程名稱(關鍵字)"
                If String.IsNullOrEmpty(CourseNameTextBox.Text.Trim) Then Msg &= "課程名稱請填寫" & vbNewLine
            Case "教師姓名(關鍵字)"
                If String.IsNullOrEmpty(TeacherNameTextBox.Text.Trim) Then Msg &= "教師姓名請填寫" & vbNewLine
            Case "授課語言(國語除外)"
                If TeachLangComboBox.SelectedIndex = -1 Then Msg &= "授課語言未選" & vbNewLine
            Case "校際通識", "通識教育中心", "體育室", "軍訓室", "師資培育中心"
                Exit Select
            Case Else
                Msg &= "程序錯誤"
        End Select '檢查資料是否有誤
        Try
            If Msg = String.Empty Then Loading(CourseWebBrowser, StatusLabel, CourseUrl, , CloseQueryForm:=True, ShowFinish:=False) '回首頁
            If Msg = String.Empty AndAlso CheckPage() Then '是否有錯誤訊息 + 
                Select Case Query.Text
                    Case "校際通識", "通識教育中心", "體育室", "軍訓室"
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("興趣"), Query.Text)
                        DoButtonClick(CourseWebBrowser, ButtonName("興趣"))
                    Case "師資培育中心"
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("教育"), Query.Text)
                        DoButtonClick(CourseWebBrowser, ButtonName("教育"))
                    Case "研究所"
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("研究所"), ResearchInstituteComboBox.SelectedItem)
                        DoButtonClick(CourseWebBrowser, ButtonName("研究所"))
                    Case "學制"
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("學制"), AcademicComboBox.SelectedItem)
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("系別"), DepartmentComboBox.SelectedItem)
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("年級"), GradeComboBox.SelectedItem)
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("班別"), ClassComboBox.SelectedItem)
                        DoButtonClick(CourseWebBrowser, ButtonName("大學和二技"))
                    Case "以時間查詢"
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("星期"), WeekComboBox.SelectedItem)
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("節次"), PeriodComboBox.SelectedItem)
                        DoButtonClick(CourseWebBrowser, ButtonName("時間"))
                    Case "課程名稱(關鍵字)"
                        DoTextBoxWrite(CourseWebBrowser, TextBoxName("課程名稱(關鍵字)"), CourseNameTextBox.Text)
                        DoButtonClick(CourseWebBrowser, ButtonName("課程"))
                    Case "教師姓名(關鍵字)"
                        DoTextBoxWrite(CourseWebBrowser, TextBoxName("教師姓名(關鍵字)"), TeacherNameTextBox.Text)
                        DoButtonClick(CourseWebBrowser, ButtonName("教師姓名"))
                    Case "授課語言(國語除外)"
                        DoComboBoxSelect(CourseWebBrowser, ComboBoxName("授課語言"), TeachLangComboBox.SelectedItem)
                        DoButtonClick(CourseWebBrowser, ButtonName("授課語言"))
                End Select '執行網頁動作
                '開始抓資料
                GetDataTimer.Start()
            Else
                HeadTabControl.Enabled = True
                WorkingToolStripProgressBar.Visible = False '隱藏進度條
                WorkingToolStripProgressBar.Style = ProgressBarStyle.Continuous
                '錯誤訊息
                If Msg <> String.Empty Then MessageBox.Show(Msg, "必要資訊", MessageBoxButtons.OK, MessageBoxIcon.Exclamation) '錯誤訊息顯示
            End If
        Catch ex As Exception
            FieldUpdate = False
            CheckPage()
        End Try
    End Sub '查詢資料並開始擷取

    Private Sub GetDataTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetDataTimer.Tick
        sender.Stop()
        '確定網頁是否重新載入
        Loading(CourseWebBrowser, StatusLabel, Nothing, CheckRepeat:=True, CloseQueryForm:=True)
        Dim AddRowResult As Boolean = ConvertHtmlTableToDataGridView(ResultDataGridView, KeepData, CourseWebBrowser, _
                                       ComputeValue, _
                                       MultiValue)
        If StartPage = 0 OrElse EndPage = 0 Then
        Else
            StatusLabel.Text = "瀏覽中..." & _
                                IIf(TermTitle Is Nothing OrElse String.IsNullOrEmpty(TermTitle.Trim), String.Empty, TermTitle & "→") & _
                                IIf(TermTitle Is Nothing OrElse String.IsNullOrEmpty(QueryTitle), String.Empty, QueryTitle & "→") & _
                                IIf(StartPage = -1 OrElse EndPage = -1, String.Empty, "第" & StartPage & "/" & EndPage & "頁")
            WorkingToolStripProgressBar.Style = ProgressBarStyle.Blocks '顯示進度條
            WorkingToolStripProgressBar.Value = (StartPage / EndPage) * 100
        End If
        If Not AddRowResult Then
            MsgBox("抱歉，此次查詢找不到課程資料。請重試")
        ElseIf EndPage - StartPage = 0 Then
            MsgBox("工作結束")
            HeadTabControl.SelectedIndex = 1 '切換至 篩選畫面
            FootTabControl.SelectedIndex = 0 '切換至 選課資料
        End If

        CheckPage()
        If EndPage - StartPage = 0 OrElse StopProcessing Then
            If PageIsAvailable Then
                HeadTabControl.Enabled = True
                StatusLabel.Text = "完成"
            Else
                StatusLabel.Text = "選課資料擷取錯誤，請按[嘗試更新]按鈕重試"
            End If
            CourseDataModule.Initialize()
            WorkingToolStripProgressBar.Style = ProgressBarStyle.Marquee '隱藏進度條
            WorkingToolStripProgressBar.Visible = False
        Else '下一個網際網路位址
            NextUri = New System.Uri(CourseWebBrowser.Url.AbsoluteUri & "?page=" & (StartPage + 1))
            For Each i As HtmlElement In CourseWebBrowser.Document.GetElementsByTagName("a")
                Dim HtmlUri As String = i.GetAttribute("href")
                Dim RequestUri As String = CourseWebBrowser.Url.Scheme & "://" & _
                CourseWebBrowser.Url.Host & CourseWebBrowser.Url.AbsolutePath & "?page=" & (StartPage + 1)
                If HtmlUri = RequestUri Then
                    i.DomElement.Click()
                    Exit For
                End If
            Next
            KeepData = True
            sender.Start()
        End If
    End Sub ' 抓選課資料

    Private Sub KeywordFilterRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        KeywordFilterListBox.Enabled = KeywordFilterRadioButton.Checked
        KeywordFilterButton.Enabled = KeywordFilterRadioButton.Checked
        ChooseValueCheckedListBox.Enabled = ChooseValueRadioButton.Checked
    End Sub '處理關鍵字的

    Private Sub TryUpdateButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TryUpdateButton.Click
        CourseDataModule.Initialize()
        ShowLoading.Show() '顯示進度視窗
        Dim LastButtonText As String = sender.Text '上次的文字
        '關閉相關控制項，避免重複執行
        HeadTabControl.Enabled = False
        TryUpdateButton.Enabled = False
        TryUpdateButton.Text = "請稍後..."
        If Loading(CourseWebBrowser, StatusLabel, CourseUrl) Then GetQueryComboBoxItem(1, 8) '更新 查詢相關控制項
        '啟用相關控制項
        TryUpdateButton.Text = LastButtonText
        ShowLoading.Close() '關閉進度視窗
        If CheckPage() Then MsgBox("欄位更新成功", MsgBoxStyle.Information, "嘗試更新結果")
        TryUpdateButton.Enabled = True
        HeadTabControl.Enabled = PageIsAvailable
    End Sub '更新查詢控制項

    '自動調整欄位
    Private Sub AutoSizeButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoSizeButton.Click
        ResultDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
        ResultDataGridView2.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
    End Sub '自動調整大小

    Private Sub ControlListBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ControlListBox.SelectedIndexChanged
        If ControlListBox.SelectedItem Is Nothing Then Exit Sub
        '初始化
        KeywordFilterListBox.Items.Clear() '清除
        '[隱藏]ChooseValueCheckedListBox.Items.Clear() '清除
        Dim FilterControl As FilterClass = ManagerFilter(ControlListBox.SelectedItem, AddIfNothing:=False)
        '找到該欄
        Dim ColumnIndex As Byte = Nothing '欄位Index
        If ResultDataGridView.Rows.Count Then
            For i As Byte = 0 To ResultDataGridView.Columns.Count - 1
                If ResultDataGridView.Columns(i).HeaderText = ControlListBox.SelectedItem Then
                    ColumnIndex = i
                    Exit For
                End If
            Next
            KeywordSplitContainer.Panel2.Enabled = True '打開右邊面板
        Else
            MsgBox("請先查詢選課資料")
            Exit Sub
        End If

        Try 'Choose Value
            If FilterControl IsNot Nothing Then KeywordFilterListBox.Items.AddRange(FilterControl.Keyword) 'Keyword
            For Each i As DataGridViewRow In ResultDataGridView.Rows
                If ChooseValueCheckedListBox.Items.IndexOf(i.Cells(ColumnIndex).Value) = -1 Then
                    ChooseValueCheckedListBox.Items.Add(i.Cells(ColumnIndex).Value)
                End If '不重複的增加方式
            Next
            '[隱藏]For Each i As String In FilterControl.Value
            '[隱藏]    If ChooseValueCheckedListBox.Items.IndexOf(i) = ListBox.NoMatches Then
            '[隱藏]        FilterControl.Remove(FilterEnum.Value, i)
            '[隱藏]    Else
            '[隱藏]        ChooseValueCheckedListBox.SetItemChecked(ChooseValueCheckedListBox.FindStringExact(i), True)
            '[隱藏]    End If
            '[隱藏]Next
        Catch ex As Exception

        End Try
    End Sub '篩選控制項

    Private Sub CompareIntCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompareIntCheckBox.CheckedChanged
        CompareIntGroupBox.Enabled = sender.Checked
        If sender.Checked = CheckState.Checked Then

        End If
    End Sub '數值比較

    Private Sub KeywordFilterButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeywordFilterButton.Click
        ManagerFilter(ControlListBox.SelectedItem, AddIfNothing:=True)
        Dim Form As New InputKeywordForm(ControlListBox.SelectedItem)
        Form.ShowDialog()
    End Sub '開啟關鍵字視窗

    Private Sub ResultDataGridView_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles ResultDataGridView.CellContentClick, ResultDataGridView.CellContentDoubleClick
        ResultDataGridView.CommitEdit(DataGridViewDataErrorContexts.CurrentCellChange) '確認使用者輸入
        Try
            If ResultDataGridView.SelectedCells.Item(0).OwningColumn.HeaderText = "課程大綱" Then
                'If MsgBox("是否開啟課程網頁？", MsgBoxStyle.YesNo, "確認") = MsgBoxResult.Yes Then
                HeadTabControl.SelectedIndex = 2
                '                CourseDetailTabPage.Show()
                CourseDetailWebBrowser.Navigate(New System.Uri(ResultDataGridView(17, ResultDataGridView.SelectedCells.Item(0).RowIndex).Value))
                'System.Diagnostics.Process.Start( _
                '     ResultDataGridView(17, ResultDataGridView.SelectedCells.Item(0).RowIndex).Value)
                'End If
            ElseIf ResultDataGridView.SelectedCells.Item(0).OwningColumn.HeaderText = "選取" Then '使用者藉由CheckBox改變Cell外觀
                Dim Index As Integer = ResultDataGridView.SelectedCells.Item(0).RowIndex
                Dim CellStyle As New DataGridViewCellStyle
                If ResultDataGridView.SelectedCells.Item(0).Value = "False" Then '未選取
                    CellStyle = ResultDataGridView.DefaultCellStyle
                Else '選取
                    CellStyle.Font = New Font(ResultDataGridView.Font, FontStyle.Bold)
                    CellStyle.BackColor = Color.Orange
                End If
                For Each i As DataGridViewCell In ResultDataGridView.Rows(Index).Cells
                    i.Style = CellStyle
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub '連線到課程網頁 + 使用者選定的課程

    'Private Sub ChooseValueCheckedListBox_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs)
    'Dim FilterControl As FilterClass = ManagerFilter(ControlListBox.SelectedItem, AddIfNothing:=True)
    '    If e.NewValue = CheckState.Checked Then
    '        FilterControl.Value = ChooseValueCheckedListBox.Items.Item(e.Index)
    '    ElseIf e.NewValue = CheckState.Unchecked Then
    '        FilterControl.Remove(FilterEnum.Value, ChooseValueCheckedListBox.Items.Item(e.Index))
    '    End If
    'End Sub 'ChooseValue更新'隱藏

    Function ManagerFilter(ByVal ControlName As String, ByVal AddIfNothing As Boolean) As FilterClass
        Dim FilterControl As FilterClass = FilterModule.Filter(ControlListBox.SelectedItem) 'Keyword
        If FilterControl Is Nothing AndAlso AddIfNothing Then FilterControl = FilterModule.Add(ControlListBox.SelectedItem)
        Return FilterControl
    End Function '簡易管理Filter新增

    'Private Sub FilterIntRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    'LessThanRadioButton.CheckedChanged, _
    'LessThanOrEqualToRadioButton.CheckedChanged, _
    'EqualToRadioButton.CheckedChanged, _
    'GreaterThanRadioButton.CheckedChanged, _
    'GreaterThanOrEqualToRadioButton.CheckedChanged
    '
    'End Sub

    Private Sub ClearFilterButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearFilterButton.Click
        FilterModule.Format()
    End Sub '清除篩選條件

    Private Sub DoFilterButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DoFilterButton.Click
        'DataColumn.Expression 屬性
        Dim [DataTable] As DataTable = DataGridViewToDataTable(ResultDataGridView)
        If DataGridViewFilter(DataGridViewToDataTable(ResultDataGridView), _
                            ResultDataGridView2, _
                            GetFilterExpressionString()) Then

            '處理課堂不一致問題
            Dim CheckClassId As New HybridDictionary '選取多時段的選課
            Dim CheckClassIdIndex As New NameValueCollection '欲刪除的位置
            'ResultDataGrieView2 的
            For Each i As System.Windows.Forms.DataGridViewRow In ResultDataGridView2.Rows
                If (Not String.IsNullOrEmpty(i.Cells(1).Value)) Then
                    If CheckClassId.Contains(i.Cells(2).Value) Then
                        CheckClassId.Item(i.Cells(2).Value) += 1
                        CheckClassIdIndex.Add(i.Cells(2).Value, i.Index)
                    Else
                        CheckClassId.Add(i.Cells(2).Value, 1)
                        CheckClassIdIndex.Add(i.Cells(2).Value, i.Index)
                    End If
                End If
            Next
            'ResultDataGrieView 的
            Dim Var As String = Nothing
            Dim LinqQuery = From ClassName2 In [DataTable].AsEnumerable _
                Let ClassName1 = Var _
                Where (ClassName2.Field(Of String)("課號")) = ClassName1 _
                Select ClassName2
            Dim RemoveAt As New ArrayList
            For Each Key As String In CheckClassId.Keys
                Var = Key
                If CheckClassId.Item(Key) <> LinqQuery.Count Then
                    For Each Index As String In CheckClassIdIndex(Key)
                        'ResultDataGridView2.Rows.RemoveAt(Index)
                        RemoveAt.Add(Index)
                    Next
                End If
            Next
            '由後往前刪除(由前往後刪會有問題-後面的row會取代之前被刪除的Index)
            Dim LinqQuery2 = From Num As Int32 In RemoveAt Order By Num Descending
            For Each Index As Int32 In LinqQuery2
                ResultDataGridView2.Rows.RemoveAt(Index)
            Next
        End If

    End Sub '篩選

    Private Sub WeekAndPeriodDataGridView_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles WeekAndPeriodDataGridView.CellContentClick
        WeekAndPeriodDataGridView.CommitEdit(DataGridViewDataErrorContexts.Commit) '確認使用者輸入
    End Sub '星期＆節次

    Private Sub ResultDataGridView2_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles ResultDataGridView2.CellClick
        If ResultDataGridView2.SelectedCells.Count = 1 AndAlso _
           ResultDataGridView2.Columns(e.ColumnIndex).HeaderText = "課程大綱" AndAlso _
           ResultDataGridView2(e.ColumnIndex, e.RowIndex).Value = "[進入]" Then
            HeadTabControl.SelectedIndex = 2
            CourseDetailWebBrowser.Navigate(New System.Uri(ResultDataGridView2(17, ResultDataGridView2.SelectedCells.Item(0).RowIndex).Value))
        End If

    End Sub '連到課程內容

    Private Sub Main_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If MsgBox("確定要離開了嗎？", MsgBoxStyle.YesNo, "確認") = MsgBoxResult.No Then e.Cancel = True
    End Sub
End Class
