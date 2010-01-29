Imports System.Collections.Specialized '字串集合
Public Module BrowserFunctionModule
    Public LastHTMLData As String '上一次的Document資料
    Public TermTitle As String = Nothing '哪個學期
    Public QueryTitle As String = Nothing '查詢條件
    '檢測網頁
    Sub WebBrowserLoading(ByRef [WebBrowser] As WebBrowser)
        Dim Time As New Stopwatch
        Time.Start()
        Do Until ([WebBrowser].ReadyState = WebBrowserReadyState.Complete _
          AndAlso [WebBrowser].IsBusy = False _
          OrElse Time.Elapsed.TotalSeconds > 30 _
          OrElse StopProcessing)
            Application.DoEvents()
        Loop '是否完成載入
        Time.Stop()
        If Time.Elapsed.TotalSeconds > 30 Then
            TimeOut = True
        End If
    End Sub '等WebBrowser載入完成

    Public Function Loading(ByRef [WebBrowser] As System.Windows.Forms.WebBrowser, _
                            ByRef StatusLabelText As System.Object, _
                            ByVal Uri As System.Uri, _
                            Optional ByVal CheckRepeat As Boolean = False, _
                            Optional ByVal CloseQueryForm As Boolean = False, _
                            Optional ByVal ShowFinish As Boolean = True)
        If Uri IsNot Nothing Then
            [WebBrowser].Navigate(Uri) '瀏覽網頁
            StatusLabelText.Text = "取得資料中..."
        End If
        WebBrowserLoading([WebBrowser])
        Dim Time As New Stopwatch
        Time.Start()
        If CheckRepeat Then
            Do Until Time.Elapsed.TotalSeconds > 10 OrElse LastHTMLData <> WebBrowser.DocumentText
                Application.DoEvents()
            Loop
        End If
        Time.Stop()
        LastHTMLData = WebBrowser.DocumentText

        '時間逾時
        If TimeOut Then Return False
        '參考網址http://www.blueshop.com.tw/board/show.asp?subcde=BRD200908232323136XL&fumcde=FUM20050124191756KKC
        '防止出現Script錯誤
        Dim Java As mshtml.IHTMLWindow2 = [WebBrowser].Document.Window.DomWindow
        Java.execScript("window.onerror = function(){return true;};" & _
                        "function confirm(){return true;};" & _
                        "function   alert(){};", "javascript")
        If Uri IsNot Nothing AndAlso ShowFinish Then StatusLabelText.Text = "完成" '狀態
        Dim Result As Boolean = (WebBrowser.Url = Uri) OrElse Uri Is Nothing
        Try
            Dim a As String = WebBrowser.Document.Domain
        Catch ex As Exception
            Result = False
        End Try
        PageIsAvailable = Result
        If Not CloseQueryForm Then Main.HeadTabControl.Enabled = Result '啟動、關閉控制項
        Return Result
    End Function '等待網頁讀取完成

    '網頁自動化
    Sub DoComboBoxSelect(ByRef Browser As Windows.Forms.WebBrowser, ByVal Name As String, ByVal SelectedItem As String)
        Dim Counter As UShort = 0
        Try
            For Each i As HtmlElement In Browser.Document.GetElementsByTagName("select").GetElementsByName(Name).Item(0).Children
                If i.InnerText = SelectedItem Then
                    Browser.Document.GetElementsByTagName("select").GetElementsByName(Name).Item(0).DomElement.SelectedIndex = Counter
                    Exit For
                End If
                Counter += 1
            Next
            PageIsAvailable = True
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub '自動選擇ComboBox + Throw

    Sub DoButtonClick(ByRef Browser As Windows.Forms.WebBrowser, ByVal Name As String)
        Try
            Browser.Document.GetElementsByTagName("input").GetElementsByName(Name).Item(0).DomElement.Click()
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub '自動點擊Button + Throw

    Sub DoTextBoxWrite(ByRef Browser As Windows.Forms.WebBrowser, ByVal Name As String, ByVal Str As String)
        Try
            Browser.Document.GetElementsByTagName("input").GetElementsByName(Name).Item(0).InnerText = Str
        Catch ex As Exception
            Throw New Exception '無法新增時，擲出一個錯誤
        End Try
    End Sub '自動寫入TextBox + Throw

    Sub GetComboBoxInformation(ByRef [ComboBox] As Windows.Forms.ComboBox, ByRef Browser As Windows.Forms.WebBrowser, ByVal Name As String, ByVal Start As Byte)
        [ComboBox].Items.Clear() '清除舊資料
        Try
            For i As Integer = Start To Browser.Document.GetElementsByTagName("select").GetElementsByName(Name).Item(0).Children.Count - 1
                ComboBox.Items.Add(Browser.Document.GetElementsByTagName("select").GetElementsByName(Name).Item(0).Children(i).InnerText)
            Next
        Catch ex As Exception
            Throw New Exception '無法新增時，擲出一個錯誤
        End Try
    End Sub '擷取資料 + Throw

    Public Function GetMultiValueColumnIndex(ByVal Str As String, _
                                             ByVal NoStr As String, _
                                             ByRef HtmlTableCellData As StringCollection, _
                                             ByRef TableMapHtmlCloumnIndex() As Short, _
                                             ByRef TableColumnName As StringCollection)
        Dim LinqQuery = From j In HtmlTableCellData Where j Like Str And Not j Like NoStr Select j
        Dim SplitStr() As String = Str.Split("*") '以＊分割
        Dim AddStr As New StringCollection
        Dim Count As Byte = Nothing
        For Each i As String In SplitStr '知道確切的數量
            If Not String.IsNullOrEmpty(i) Then AddStr.Add(i)
        Next
        For Each i As String In LinqQuery
            Dim Index As SByte = HtmlTableCellData.IndexOf(i)
            For Each ColumnName As String In AddStr
                TableMapHtmlCloumnIndex(TableColumnName.IndexOf(ColumnName)) = Index
                Count += 1
            Next
        Next
        Return IIf(Count - 1 < 0, 0, Count - 1) '減掉原來有的一欄
    End Function 'MultiValue取Index

    '網頁資料擷取
    Function ConvertHtmlTableToDataGridView(ByRef [DataGridView] As DataGridView, ByVal KeepData As Boolean, _
                                            ByRef [WebBrowser] As WebBrowser, Optional ByVal ComputeValue() As String = Nothing, _
                                            Optional ByVal MultiValue() As String = Nothing) As Boolean
        Try
            '定義
            Const COLUMN_NOT_FOUND = -1, COLUMN_COMPUTE = -2, TR_TEXT_COUNT = 1
            Dim MultiColumnAddition As Byte = Nothing

            '處理DataGridView表格欄名
            Dim TableColumnName As New StringCollection
            For i As Byte = 0 To [DataGridView].Columns.Count - 1
                TableColumnName.Add([DataGridView].Columns.Item(i).HeaderText)
            Next '取得DataGriwView的欄名
            Dim TableColumnCount As Byte = TableColumnName.Count '目前DataGridView的欄數

            '清除舊資料
            If Not KeepData Then [DataGridView].Rows.Clear() '執行前"清除"

            '讀取表格資料
            Dim HtmlTableCellData As New StringCollection
            Dim HtmlTableCellDataNoNewLine As New StringCollection
            Dim HtmlTableTagData As New StringCollection
            For Each i As HtmlElement In [WebBrowser].Document.All
                If i.TagName = "TR" OrElse i.TagName = "TD" Then
                    HtmlTableTagData.Add(i.TagName) '標籤
                    Dim CellData As String = IIf(i.InnerText Is Nothing, String.Empty, i.InnerText)
                    HtmlTableCellData.Add(CellData) '值
                    HtmlTableCellDataNoNewLine.Add(CellData.Replace(vbNewLine, String.Empty)) '值 沒換行
                End If
            Next
            Dim LinqQuery = From PageNum In HtmlTableCellDataNoNewLine _
                            Where PageNum Like "*第*/*頁*" OrElse PageNum Like "*第*\*頁*" _
                            Select PageNum

            '擷取第幾頁
            Dim TotalPage As String = Nothing '第*/*頁
            StartPage = COLUMN_NOT_FOUND
            EndPage = COLUMN_NOT_FOUND
            For Each i As String In LinqQuery
                TotalPage = i
                Dim SplitStr() As String = i.Split(New Char() {"第", "/", "\", "頁"})
                For StrIndex As Byte = 1 To SplitStr.Count - 1
                    If IsNumeric(SplitStr(StrIndex)) AndAlso (Not String.IsNullOrEmpty(SplitStr(StrIndex))) AndAlso StartPage = COLUMN_NOT_FOUND Then
                        StartPage = CInt(SplitStr(StrIndex))
                    ElseIf IsNumeric(SplitStr(StrIndex)) AndAlso (Not String.IsNullOrEmpty(SplitStr(StrIndex))) AndAlso StartPage <> COLUMN_NOT_FOUND Then
                        EndPage = CInt(SplitStr(StrIndex))
                        Exit For
                    End If
                Next
                If EndPage <> COLUMN_NOT_FOUND OrElse StartPage <> COLUMN_NOT_FOUND Then
                    Exit For
                End If
            Next
            If TotalPage Is Nothing Then Return False '沒有第幾頁代表查無資料

            '標題
            TermTitle = HtmlTableCellData(1) '第幾學期
            QueryTitle = HtmlTableCellData(3) '查詢條件Title

            '找開始 結束
            Dim StartCellIndex As Int32 = Nothing
            Dim EndCellIndex As Int32 = Nothing

            '處理表格Table To Html對應 + 開始欄Index ~ 結束欄Index
            Dim TableMapHtmlCloumnIndex(TableColumnCount - 1) As Short
            For i As Byte = 0 To TableColumnCount - 1
                TableMapHtmlCloumnIndex(i) = HtmlTableCellDataNoNewLine.IndexOf(TableColumnName(i))
                If TableMapHtmlCloumnIndex(i) <> COLUMN_NOT_FOUND Then
                    EndCellIndex = TableMapHtmlCloumnIndex(i)
                ElseIf StartCellIndex = Nothing AndAlso TableMapHtmlCloumnIndex(i) <> COLUMN_NOT_FOUND Then
                    StartCellIndex = TableMapHtmlCloumnIndex(i)
                End If
            Next '欄位對應表

            For Each i As String In ComputeValue
                TableMapHtmlCloumnIndex(TableColumnName.IndexOf(i)) = COLUMN_COMPUTE
            Next '-2代表 自動計算的值

            '*星期*節次*教室*
            MultiColumnAddition += GetMultiValueColumnIndex(MultiValue(0), MultiValue(1), HtmlTableCellDataNoNewLine, TableMapHtmlCloumnIndex, TableColumnName)
            '*選上人數*
            MultiColumnAddition += GetMultiValueColumnIndex(MultiValue(1), MultiValue(0), HtmlTableCellDataNoNewLine, TableMapHtmlCloumnIndex, TableColumnName)

            'HTML欄數 + 開始位置 + 結束位置 + 有幾欄
            Dim HtmlTableColumnCount As Int32 = Nothing
            For ColumnIndex As Byte = 0 To TableMapHtmlCloumnIndex.Count - 1
                If TableMapHtmlCloumnIndex(ColumnIndex) >= 0 Then
                    HtmlTableColumnCount += 1
                End If
            Next
            HtmlTableColumnCount -= MultiColumnAddition

            StartCellIndex = EndCellIndex + 1
            EndCellIndex = HtmlTableCellDataNoNewLine.IndexOf(TotalPage) - 1

            Dim HtmlTableRowCount As Int16 = (EndCellIndex - StartCellIndex + 1) / (HtmlTableColumnCount + TR_TEXT_COUNT)

            '
            '解析選課資料
            '
            Dim NewRowData(TableColumnCount) As String '待新增的ROW
            Dim WeekPeriodLocation(60) As String '待新增的星期 節次 地點
            Dim Repeat As Byte = 0 '重複輸入該課程的不同時段
            Dim IndexSelect As Byte = TableColumnName.IndexOf("選取") '              00
            Dim IndexTime As Byte = TableColumnName.IndexOf("時段") '                01
            Dim IndexCourse As Byte = TableColumnName.IndexOf("課程名稱") '          04
            Dim IndexHasCourseLink As Byte = TableColumnName.IndexOf("課程大綱") '   05
            Dim IndexWeek As Byte = TableColumnName.IndexOf("星期") '                10
            Dim IndexPeriod As Short = TableColumnName.IndexOf("節次") '             11
            Dim IndexClassRoom As Byte = TableColumnName.IndexOf("教室") '           12
            Dim IndexPopularityLimit As Short = TableColumnName.IndexOf("人數限制") '13
            Dim IndexPopularitySelected As Byte = TableColumnName.IndexOf("選上人數") '14
            Dim IndexRemainingPlaces As Byte = TableColumnName.IndexOf("剩下名額") ' 15
            Dim IndexCourseLink As Byte = TableColumnName.IndexOf("課程連結") '      17

            Dim TotalCourseLink As New StringCollection
            For Each i As HtmlElement In [WebBrowser].Document.GetElementsByTagName("a")
                Dim TagA As String = i.GetAttribute("href").StartsWith("http://course.chu.edu.tw/outline.asp?courseid=")
                If i.GetAttribute("href").StartsWith("http://course.chu.edu.tw/outline.asp?courseid=") Then
                    TotalCourseLink.Add(i.GetAttribute("href"))
                End If
            Next '所有課程網址
            Dim CourseLinkIndex As Byte = Nothing
            Dim HasCourseLink As Boolean = Nothing
            For RowIndex As Byte = 0 To HtmlTableRowCount - 1
                For ColumnIndex As Byte = 0 To TableColumnCount - 1
                    If TableMapHtmlCloumnIndex(ColumnIndex) = COLUMN_NOT_FOUND Then '網頁上沒有這個列
                        Continue For
                        '------------------------------------------
                    ElseIf TableMapHtmlCloumnIndex(ColumnIndex) = COLUMN_COMPUTE Then '程式產生出來的
                        If IndexSelect = ColumnIndex Then '選取
                            NewRowData(ColumnIndex) = False
                        ElseIf IndexTime = ColumnIndex Then '時段
                            NewRowData(ColumnIndex) = String.Empty
                        ElseIf IndexRemainingPlaces = ColumnIndex Then '剩下名額
                            NewRowData(ColumnIndex) = _
                            IIf(IsNumeric(NewRowData(IndexPopularityLimit)) _
                            AndAlso IsNumeric(NewRowData(IndexPopularitySelected)) _
                            AndAlso NewRowData(IndexPopularityLimit) - NewRowData(IndexPopularitySelected) > 0, _
                            NewRowData(IndexPopularityLimit) - NewRowData(IndexPopularitySelected), _
                            0)
                        ElseIf IndexHasCourseLink = ColumnIndex Then '[課程網址]-進入
                            NewRowData(ColumnIndex) = Nothing
                        ElseIf IndexCourseLink = ColumnIndex Then '[課程網址]-連結
                            If HasCourseLink Then '如果有連結
                                NewRowData(ColumnIndex) = TotalCourseLink(CourseLinkIndex)
                                NewRowData(IndexHasCourseLink) = "[進入]"
                                CourseLinkIndex += 1
                                HasCourseLink = False
                            Else
                                NewRowData(IndexHasCourseLink) = String.Empty
                            End If
                        End If
                        '------------------------------------------
                    ElseIf IndexWeek = ColumnIndex Then '星期 節次 教室
                        Dim Symbol() As Char = {"(", ")", "【", "】"}
                        Dim Seperator As Byte = 0
                        Dim TempStr As String = _
                        HtmlTableCellData((RowIndex + 1) * (HtmlTableColumnCount + TR_TEXT_COUNT) + _
                                          TableMapHtmlCloumnIndex(ColumnIndex))
                        If TempStr.IndexOf("無資料") >= 0 OrElse String.IsNullOrEmpty(TempStr) Then
                            WeekPeriodLocation(0) = "無資料"
                            WeekPeriodLocation(1) = "無資料"
                            WeekPeriodLocation(2) = "無資料"
                            Repeat = 1
                            Continue For
                        Else
                            Dim WeekPeriodLocationSubStr() As String = TempStr.Split(New Char() {vbNewLine}) '暫存Str
                            For StrNum As Byte = 0 To WeekPeriodLocationSubStr.Count - 1
                                WeekPeriodLocationSubStr(StrNum) = WeekPeriodLocationSubStr(StrNum).Trim
                                For CharNum As Byte = 0 To WeekPeriodLocationSubStr(StrNum).Length - 1
                                    Select Case WeekPeriodLocationSubStr(StrNum).Chars(CharNum)
                                        Case Symbol(0)
                                            Seperator = IIf(Seperator = 2, 2, 0)
                                            Continue For
                                        Case Symbol(1)
                                            Seperator = 1
                                        Case Symbol(2)
                                            Seperator = 2
                                        Case Symbol(3)
                                            Seperator = 3
                                        Case Else
                                            If Seperator < 3 Then '0-2
                                                WeekPeriodLocation(Seperator + Repeat * 3) &= WeekPeriodLocationSubStr(StrNum).Chars(CharNum)
                                            End If
                                    End Select
                                Next
                                Repeat += 1
                            Next
                        End If
                        ColumnIndex += MultiColumnAddition
                        '------------------------------------------
                    ElseIf ColumnIndex = TableColumnCount - 1 Then '最後一個
                        NewRowData(ColumnIndex) = HtmlTableCellData((RowIndex + 1) * (HtmlTableColumnCount + TR_TEXT_COUNT) + TableMapHtmlCloumnIndex(ColumnIndex))
                        For i As Byte = 1 To Repeat
                            '讀取多時段 時間和教室
                            NewRowData(IndexWeek) = WeekPeriodLocation((i - 1) * 3 + 0)
                            NewRowData(IndexPeriod) = WeekPeriodLocation((i - 1) * 3 + 1)
                            NewRowData(IndexClassRoom) = WeekPeriodLocation((i - 1) * 3 + 2)
                            NewRowData(IndexTime) = IIf(Repeat > 1, "(" & i & ")", String.Empty) '同課程
                            [DataGridView].Rows.Add(NewRowData) '新增列
                        Next
                        If Repeat = 0 Then
                            [DataGridView].Rows.Add(NewRowData)
                        End If
                        '重設內容
                        ReDim NewRowData(TableColumnCount)
                        ReDim WeekPeriodLocation(15)
                        Repeat = 0
                        '------------------------------------------
                    ElseIf IndexCourse = ColumnIndex Then '課程名稱
                        Dim SplitStr() As String = HtmlTableCellData((RowIndex + 1) * (HtmlTableColumnCount + TR_TEXT_COUNT) + TableMapHtmlCloumnIndex(ColumnIndex)).Split(New Char() {vbNewLine, "~"})
                        SplitStr(0) = SplitStr(0).Replace("（", "(")
                        SplitStr(0) = SplitStr(0).Replace("）", ")").Trim
                        HasCourseLink = IIf(HtmlTableCellData((RowIndex + 1) * (HtmlTableColumnCount + TR_TEXT_COUNT) + TableMapHtmlCloumnIndex(ColumnIndex)).IndexOf("課程大綱") <> -1, True, False)
                        NewRowData(ColumnIndex) = SplitStr(0)
                        '------------------------------------------
                    Else '其它
                        NewRowData(ColumnIndex) = HtmlTableCellData((RowIndex + 1) * (HtmlTableColumnCount + TR_TEXT_COUNT) + TableMapHtmlCloumnIndex(ColumnIndex))
                    End If
                Next
            Next
            '------------------------------------------
            [DataGridView].AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function '網頁Table轉成DataGridView + Throw

    Function DataGridViewToDataTable(ByRef [DataGridView] As DataGridView) As DataTable
        'http://social.msdn.microsoft.com/Forums/zh-TW/242/thread/0fe84c5c-6715-4f2c-af7d-40b47ff5f359
        Dim [DataTable] As New DataTable
        For i As Integer = 0 To [DataGridView].ColumnCount - 1
            Dim [DataColumn] = New DataColumn
            [DataColumn].ColumnName = [DataGridView].Columns(i).HeaderText.ToString
            [DataTable].Columns.Add([DataColumn])
        Next '取得標題
        For i As Integer = 0 To [DataGridView].RowCount - 1
            Dim [DataRow] As DataRow = [DataTable].NewRow
            For j As Integer = 0 To [DataGridView].ColumnCount - 1
                [DataRow](j) = [DataGridView].Rows(i).Cells(j).Value
            Next
            [DataTable].Rows.Add([DataRow])
        Next
        Return [DataTable]
    End Function 'DataGridView轉成DataTable

    '網頁錯誤訊息
    Function CheckPage()
        If Not PageIsAvailable Then
            If TimeOut Then
                TimeOut = Nothing
                MsgBox("等候時間逾時", MsgBoxStyle.Exclamation, "逾時")
            ElseIf StopProcessing Then
                StopProcessing = Nothing
                MsgBox("使用者取消工作", MsgBoxStyle.Exclamation, "已取消")
            ElseIf FieldUpdate Then
                FieldUpdate = Nothing
                MsgBox("欄位更新錯誤，請重試", MsgBoxStyle.Critical, "更新")
            ElseIf Not PageIsAvailable Then
                MsgBox("擷取資料失敗!" & vbNewLine & _
                        "請檢查" & vbNewLine & _
                        "1.電腦是否已連上網路" & vbNewLine & _
                        "2.學校主機是否能連線" & vbNewLine & _
                        "3.您的IE是否處於離線狀態(在IE功能表的 檔案->離線工作)" & vbNewLine & _
                        "-確認後，如果持續出現此訊息，請嘗試重新啟動此程式，謝謝。-", MsgBoxStyle.Critical, "連線錯誤")
            End If
        End If
        Return PageIsAvailable '回傳結果
    End Function '檢查是否有誤，並顯示錯誤

End Module
