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
End Module
