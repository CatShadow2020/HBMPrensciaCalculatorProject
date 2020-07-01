Public Class ResultClass
    Private m_ResultList() As String
    Private m_NResults As Integer

    Public Sub New(nSz As Integer)
        m_NResults = 0
        ReDim m_ResultList(nSz)
    End Sub

    Public Function getResult() As String
        Return m_ResultList(0)
    End Function

    Public Function getPreviousResult(index As Integer) As String
        If index < 1 Or index > UBound(m_ResultList) Then
            '
            ' Technically, index = 0 works too, 
            ' but in the function specification index = [1,10]
            '
            Return Nothing
        End If
        Return m_ResultList(index)
    End Function

    Protected Sub StoreResult(s As String)
        Dim i As Integer
        Dim nLastIndex As Integer

        If m_NResults > 0 Then
            If m_NResults = m_ResultList.Length Then
                nLastIndex = m_NResults - 1
            Else
                nLastIndex = m_NResults
            End If

            For i = nLastIndex - 1 To 0 Step -1
                m_ResultList(i + 1) = m_ResultList(i)
            Next i
        End If

        m_ResultList(0) = s

        If m_NResults < m_ResultList.Length Then
            m_NResults += 1
        End If

    End Sub

End Class
