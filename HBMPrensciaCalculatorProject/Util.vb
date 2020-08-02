Public Module Util
    Public Function SplitCalculatorString(s As String) As String()
        Dim sOperatorLis() As Char = {"*", "/", "+", "-"}
        Dim numberList() As String
        numberList = s.Split(sOperatorLis)
        Return numberList
    End Function

    Public Function CheckLastNumberForDot(s As String) As Boolean
        Dim numberList() As String
        numberList = SplitCalculatorString(s)
        If numberList.Count = 0 Then Return False
        If numberList.Last.IndexOf(".") = -1 Then Return False
        Return True
    End Function

    Public Function CheckParentheses(s As String) As Boolean
        Dim nLeft, nRight As Integer
        Dim i As Integer
        nLeft = 0
        nRight = 0

        For i = 0 To s.Length - 1
            If s(i) = "(" Then
                nLeft += 1
                Continue For
            End If

            If s(i) = ")" Then
                nRight += 1
                If nRight > nLeft Then Return False
                Continue For
            End If
        Next

        If (nLeft <> nRight) Then Return False

        Return True

    End Function
End Module
