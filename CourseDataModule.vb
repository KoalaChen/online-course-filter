Module CourseDataModule
    '頁數
    Private StartPageField As Int32 = Nothing '開始
    Private EndPageField As Int32 = Nothing '結束
    '頁面管理
    Private PageIsAvailableField As Boolean = Nothing '頁面是否有效
    Public KeepData As Boolean = False '是否留存DataGridView資料
    Public StopProcessing As Boolean = False '使用這個值可以停止程式擷取資料
    Public TimeOut As Boolean = False '時間逾時
    Public FieldUpdate As Boolean = False '網頁控制項內容是否更新
    Public NextUri As System.Uri = CourseUrl '下一個網址
    Public AcademicComboBoxIndex As Short = -1 '學制控制項
    '課程網址
    Public CourseUrlField As New System.Uri("http://course.chu.edu.tw/") '課程網址
    '網頁上定義的控制項名稱
    ReadOnly Property ComboBoxName(ByVal RequestName As String) As String
        Get
            Select Case RequestName
                Case "興趣"
                    Return "order_select"
                Case "教育"
                    Return "edu_select"
                Case "研究所"
                    Return "master_select"
                Case "學制"
                    Return "college_select"
                Case "系別"
                    Return "dept_select"
                Case "年級"
                    Return "grade_select"
                Case "班別"
                    Return "class_select"
                Case "星期"
                    Return "by_week"
                Case "節次"
                    Return "by_period"
                Case "授課語言"
                    Return "language_select"
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property '網頁上ComboBox的名稱

    ReadOnly Property ButtonName(ByVal RequestName As String) As String
        Get
            Select Case RequestName
                Case "興趣"
                    Return "order_course"
                Case "教育"
                    Return "edu_course"
                Case "大學和二技"
                    Return "college_course"
                Case "研究所"
                    Return "master_course"
                Case "時間"
                    Return "time_course"
                Case "課程"
                    Return "text_course"
                Case "教師姓名"
                    Return "teacher_name"
                Case "授課語言"
                    Return "language_course"
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property '網頁上Button的名稱

    ReadOnly Property TextBoxName(ByVal RequestName As String) As String
        Get
            Select Case RequestName
                Case "課程名稱(關鍵字)"
                    Return "keyword"
                Case "教師姓名(關鍵字)"
                    Return "keyword_t"
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property '網頁上TextBox的名稱

    '網頁屬性
    Property PageIsAvailable() As Boolean
        Get
            Return PageIsAvailableField
        End Get
        Set(ByVal value As Boolean)
            PageIsAvailableField = value
        End Set
    End Property '頁面是否有效

    Property StartPage() As Integer
        Get
            Return StartPageField
        End Get
        Set(ByVal value As Integer)
            StartPageField = value
        End Set
    End Property '開始

    Property EndPage() As Integer
        Get
            Return EndPageField
        End Get
        Set(ByVal value As Integer)
            EndPageField = value
        End Set
    End Property '結束

    ReadOnly Property CourseUrl() As System.Uri
        Get
            Return CourseUrlField
        End Get
    End Property '課程網址

    Sub Initialize() '網頁參數初始化
        StartPage = Nothing
        EndPage = Nothing
        PageIsAvailable = Nothing
        KeepData = Nothing
        StopProcessing = Nothing
        TimeOut = Nothing
        FieldUpdate = Nothing
        AcademicComboBoxIndex = -1
        Main.ReFreshDataLabel.Visible = False
        Main.AcademicComboBox.Enabled = True
        Main.DepartmentComboBox.Enabled = True
    End Sub
End Module
