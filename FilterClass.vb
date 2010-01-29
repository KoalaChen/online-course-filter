Imports System.Collections.Specialized '字串陣列
Public Class FilterClass
    Private KeywordField As New StringCollection
    'Private ValueField As New StringCollection
    Sub New()

    End Sub

    Sub Remove(ByVal Type As FilterEnum, ByVal Str As String)
        If Type = FilterEnum.Keyword Then
            KeywordField.Remove(Str)
            'ElseIf Type = FilterEnum.Value Then
            '    ValueField.Remove(Str)
        End If
    End Sub

    Property Keyword(Optional ByVal Index As Integer = -1)  '關鍵字
        Get
            If Index = -1 Then
                If KeywordField.Count = 0 Then
                    Return Nothing
                Else
                    Dim Result(KeywordField.Count - 1) As String
                    For i As Int16 = 0 To KeywordField.Count - 1
                        Result(i) = KeywordField(i)
                    Next
                    Return Result
                End If
            Else
                Return KeywordField.Item(Index)
            End If
        End Get
        Set(ByVal value)
            If Not KeywordField.Contains(value) Then
                KeywordField.Add(value)
            End If
        End Set
    End Property
    'Property Value(Optional ByVal Index As Integer = -1)  '已出現的字
    '    Get
    '        If Index = -1 Then
    '            Dim Result(ValueField.Count - 1) As String
    '            For i As Int16 = 0 To ValueField.Count - 1
    '                Result(i) = ValueField(i)
    '            Next
    '            Return Result
    '        Else
    '            Return ValueField.Item(Index)
    '        End If
    '    End Get
    '    Set(ByVal value)
    '        If Not ValueField.Contains(value) Then
    '            ValueField.Add(value)
    '        End If
    '    End Set
    'End Property
End Class
