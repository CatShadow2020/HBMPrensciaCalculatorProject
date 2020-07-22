Module ArithmeticModule
    Public Function IsValidOperator(value As String) As Boolean
        If value <> "+" And value <> "-" And value <> "/" And value <> "*" Then
            Return False
        End If

        Return True
    End Function

    '
    ' Convert two strings to Double array(1)
    ' Function does not validate arguments
    ' It need to be done before
    '
    Private Function Parameters2Double(left As String, right As String) As Double()
        Dim lR(1) As Double
        lR(0) = Convert.ToDouble(left)
        lR(1) = Convert.ToDouble(right)
        Return lR
    End Function


    Private Function AddOperator(left As String, right As String) As String
        Dim lResult As Double
        Dim sResult As String
        Dim lR() As Double

        Try
            lR = Parameters2Double(left, right)
            lResult = lR(0) + lR(1)
            sResult = Convert.ToString(lResult)
        Catch ex As Exception
            Return ERROR_COLON + ex.ToString
        End Try

        Return sResult

    End Function

    Private Function SubstractOperator(left As String, right As String) As String
        Dim lResult As Double
        Dim sResult As String
        Dim lR() As Double

        Try
            lR = Parameters2Double(left, right)
            lResult = lR(0) - lR(1)
            sResult = Convert.ToString(lResult)
        Catch ex As Exception
            Return ERROR_COLON + ex.ToString
        End Try

        Return sResult
    End Function

    Private Function MultiplyOperator(left As String, right As String) As String
        Dim lResult As Double
        Dim sResult As String
        Dim lR() As Double

        Try
            lR = Parameters2Double(left, right)
            lResult = lR(0) * lR(1)
            sResult = Convert.ToString(lResult)
        Catch ex As Exception
            Return ERROR_COLON + ex.ToString
        End Try

        Return sResult
    End Function

    Private Function DivideOperator(left As String, right As String) As String
        Dim lResult As Double
        Dim sResult As String
        Dim lR() As Double

        Try
            lR = Parameters2Double(left, right)
            If lR(1) = 0 Then
                Return ZERO_DIVIDE
            End If
            lResult = lR(0) / lR(1)
            sResult = Convert.ToString(lResult)
        Catch ex As Exception
            Return ERROR_COLON + ex.ToString
        End Try

        Return sResult
    End Function

    '
    ' This is calculate code which used in both calculator
    ' classes
    '
    '
    Public Function calculate_(left As String, right As String, oper As String) As String
        Dim resultString As String = INTERNAL_ERROR

        If IsNumeric(left) = False Then
            resultString = LEFT_PARAMETERS_IS_INVALID
            Return resultString
        End If

        If IsNumeric(right) = False Then
            resultString = RIGHT_PARAMETERS_IS_INVALID
            Return resultString
        End If

        If IsValidOperator(oper) = False Then
            resultString = OPERATOR_IS_INVALID
            Return resultString
        End If

        Select Case oper
            Case "*"
                resultString = MultiplyOperator(left, right)
            Case "/"
                resultString = DivideOperator(left, right)
            Case "-"
                resultString = SubstractOperator(left, right)
            Case "+"
                resultString = AddOperator(left, right)
        End Select
        Return resultString

    End Function

End Module
