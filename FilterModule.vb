Imports System.Collections.Specialized
Public Module FilterModule
    Public FilterField As New HybridDictionary
    Private CourseDataView As DataView
    ReadOnly Property Filter(ByVal Key As String) As FilterClass
        Get
            Return FilterField(Key)
        End Get
    End Property '篩選查詢

    Public Function Add(ByVal Key As String, Optional ByRef Value As FilterClass = Nothing) As FilterClass
        FilterField.Add(Key, New FilterClass())
        Return Filter(Key)
    End Function '新增篩選項目

    Public Sub Remove(ByVal Key As String)
        Dim OldFilter As FilterClass = FilterField(Key)
        OldFilter = Nothing
        FilterField.Remove(Key)
    End Sub '移除篩選項目

    Public Sub Format() '清除篩選條件
        FilterField.Clear()
        Main.WeekAndPeriodDataGridView.Rows.Clear()
        Dim Period() As Char = {"1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8c", "9"c, "A"c, "B"c, "C"c, "D"c}
        For Each i As Char In Period
            Main.WeekAndPeriodDataGridView.Rows.Add(New String() {"第" & i & "節", False, False, False, False, False, False, False})
        Next
        Main.KeywordFilterListBox.Items.Clear()
    End Sub '清除篩選資料

    Public Function DataGridViewFilter(ByRef CourseDataTable As DataTable, _
                                  ByRef CourseDataGridView As DataGridView, _
                                  ByRef ExpressionString As String)
        'http://www.dotblogs.com.tw/limingstudio/archive/2009/11/12/11610.aspx

        If ExpressionString Is Nothing Then MsgBox("請選擇篩選條件！", MsgBoxStyle.Information, "課程篩選") : Return (False)
        CourseDataView = CourseDataTable.DefaultView
        Main.ResultDataGridView2.DataSource = CourseDataView '設定資料來源
        CourseDataView.RowFilter = ExpressionString
        CourseDataView.RowStateFilter = DataViewRowState.CurrentRows
        Dim Time As New Stopwatch
        Time.Start()
        Do Until Time.Elapsed.TotalSeconds > 1
            Application.DoEvents()
        Loop
        Time.Stop()
        Main.FootTabControl.SelectedIndex = 1
        Main.ResultDataGridView2.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader)
        Main.ResultDataGridView2.Columns("選取").Visible = False
        Dim CellStyle As New DataGridViewCellStyle
        CellStyle.ForeColor = Color.Blue
        CellStyle.Font = New Font(Main.ResultDataGridView2.Font, FontStyle.Underline)
        Main.ResultDataGridView2.Columns("課程大綱").CellTemplate.Style = CellStyle
        Main.ResultDataGridView2.Columns("課程連結").Visible = False
        Return True
    End Function
    'Op = Operation
    Public Function GetFilterExpressionString() As String
        Dim KeywordResult As New StringCollection '蒐集關鍵字
        Dim WeekPeriodResult As New StringCollection '蒐集可以時間
        Dim WeekPeriodNotInResult As New StringCollection '蒐集不可以時間

        Dim Result As New System.Text.StringBuilder '增加效能
        Dim Operation() As String = {"And", "Or", "Not", "Like", "In"} '運算元

        '建立 關鍵字
        For i As Integer = 0 To FilterField.Keys.Count - 1
            Dim TempResult As String = String.Empty
            Dim Key As String = FilterField.Keys(i)
            Dim FilterVar As FilterClass = FilterField(Key)
            If FilterVar.Keyword Is Nothing Then
                Continue For
            Else
                Dim TempStr() As String = FilterVar.Keyword
                For j As Integer = 0 To TempStr.Count - 1
                    TempResult &= Key & TempStr(j)  'Like
                    If Not j = TempStr.Count - 1 Then
                        TempResult &= " " & Operation(OpEnum.And) & " " 'Or
                    End If
                Next
                KeywordResult.Add(TempResult)
            End If
        Next
        '星期和節次(包含)
        Dim PeriodStr() As Char = {"1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c, "A"c, "B"c, "C"c, "D"c}
        Dim WeekStr() As Char = {"日", "一", "二", "三", "四", "五", "六"}
        For i As Integer = 0 To Main.WeekAndPeriodDataGridView.RowCount - 1
            For j As Integer = 0 To Main.WeekAndPeriodDataGridView.ColumnCount - 1
                If Main.WeekAndPeriodDataGridView(j, i).Value = "True" Then
                    WeekPeriodResult.Add(" 星期 = '" & WeekStr(j - 1) & "' " & Operation(OpEnum.And) & " 節次 " & Operation(OpEnum.Like) & " '*" & PeriodStr(i) & "*' ")
                End If
            Next
        Next
        If WeekPeriodResult.Count >= 91 Then WeekPeriodResult.Clear() '若使用者全選，則取消選擇
        '星期和節次(不包含)
        If WeekPeriodResult.Count AndAlso WeekPeriodResult.Count < 91 Then '如果使用者指定時間 AND 不包含全選
            For i As Integer = 0 To Main.WeekAndPeriodDataGridView.RowCount - 1
                For j As Integer = 1 To Main.WeekAndPeriodDataGridView.ColumnCount - 1
                    If Main.WeekAndPeriodDataGridView(j, i).Value <> "True" AndAlso j <> 0 Then
                        WeekPeriodNotInResult.Add(" 星期 = '" & WeekStr(j - 1) & "' " & Operation(OpEnum.And) & " 節次 Like '*" & PeriodStr(i) & "*' ")
                    End If
                Next
            Next
        End If
        '
        '開始合併字串
        '

        '輸出空值(如果都沒有指定的話)
        If FilterField.Keys.Count = 0 AndAlso WeekPeriodResult.Count = 0 Then Return Nothing

        '輸出 關鍵字
        Result.AppendLine(IIf(FilterField.Keys.Count > 0, "(", String.Empty))
        If KeywordResult.Count Then '如果有才執行
            For i As Integer = 0 To KeywordResult.Count - 1
                Result.AppendLine(" " & KeywordResult(i) & " ")
                If i <> KeywordResult.Count - 1 Then '一直建立 AND 運算元
                    Result.AppendLine(" " & Operation(OpEnum.And) & " ")
                End If
            Next
        End If
        Result.Append(IIf(FilterField.Keys.Count > 0, ")", String.Empty))

        '輸出 星期和節次(包含)
        If Main.reverseCheckBox.Checked AndAlso WeekPeriodNotInResult.Count > 1 Then
            Result.AppendLine(IIf((FilterField.Keys.Count <> 0 AndAlso WeekPeriodResult.Count = 0) OrElse _
                                  (FilterField.Keys.Count = 0 AndAlso WeekPeriodResult.Count <> 0), " Not ", " And Not "))
        Else
            Result.AppendLine(IIf((FilterField.Keys.Count <> 0 AndAlso WeekPeriodResult.Count = 0) OrElse _
                                  (FilterField.Keys.Count = 0 AndAlso WeekPeriodResult.Count <> 0), String.Empty, " And "))
        End If
        Result.AppendLine(IIf(WeekPeriodResult.Count = 0, String.Empty, " ( "))  'TO DO:
        For i As Integer = 0 To WeekPeriodResult.Count - 1
            Result.Append(" ( " & WeekPeriodResult(i) & " ) ")
            If i <> WeekPeriodResult.Count - 1 Then
                Result.AppendLine(" " & Operation(OpEnum.Or) & " ")
            End If
        Next
        Result.AppendLine(IIf(WeekPeriodResult.Count = 0, String.Empty, " ) "))

        Dim test As String = Result.ToString
        '輸出 星期和節次(不包含)
        If Main.reverseCheckBox.Checked Then
            Result.AppendLine(IIf( _
                            (WeekPeriodResult.Count <> 0 AndAlso WeekPeriodNotInResult.Count <> 0) _
                           , " And  (", String.Empty))
        Else
            Result.AppendLine(IIf( _
                            (WeekPeriodResult.Count <> 0 AndAlso WeekPeriodNotInResult.Count <> 0) _
                           , " And Not (", String.Empty))
        End If

        For i As Integer = 0 To WeekPeriodNotInResult.Count - 1
            Result.Append(" ( " & WeekPeriodNotInResult(i) & " ) ")
            If i <> WeekPeriodNotInResult.Count - 1 Then
                Result.AppendLine(" " & Operation(OpEnum.Or) & " ")
            End If
        Next
        Result.Append(IIf(WeekPeriodResult.Count = 0, String.Empty, ")"))
        If (WeekPeriodNotInResult.Count = 0) Then
            Dim a As Int64 = Result.Length
            Result.Chars(Result.Length - 1) = String.Empty
            Result.Chars(Result.Length - 2) = String.Empty
        End If
        Dim test2 As String = Result.ToString
        Return Result.ToString
    End Function '類似SQL語法

    Public Sub reFreshKeyword(ByRef tempKeyword As String(), ByRef [ListBox] As System.Windows.Forms.ListBox)
        For Each keyword As String In tempKeyword
            If keyword.Contains("Not Like") Then
                Dim tempStr() As String = keyword.Split(New Char() {"*"})
                [ListBox].Items.Add(tempStr(1) & "(不包含)") 'Keyword
            Else
                Dim tempStr() As String = keyword.Split(New Char() {"*"})
                [ListBox].Items.Add(tempStr(1)) 'Keyword
            End If
        Next
    End Sub
End Module
