' The Calculator Class Is a contrived example, which has a very deliberate API. Implement an alternate Class that accomplishes the same goal (simple arithmetic And storing the last 10 results) but With a less verbose API. 
'  
' This second implementation should implement a method that returns the result Of a calculation 
'  
'     calculate(equation)

Public Class CalculatorPart2
    Inherits ResultClass
    Public Sub New()
        MyBase.New(10)
    End Sub

    '
    ' Calculates equation 's' :
    ' The format of an equation is 
    ' "<number> <operator> <number2> <operator2> ... <numberN> <operatorN>", 
    ' starting And ending with a digit. 
    '
    '
    Public Sub calculate(s As String)
        Dim sOperatorLis() As Char = {"*", "/", "+", "-"}


        Dim operatorList As New List(Of String)()
        Dim numberList() As String
        Dim nPos As Integer = 0
        Dim c As Char
        Dim i As Integer
        Dim left, right, oper As String
        Dim prevRes As String

        If IsNothing(s) Then
            StoreResult(INCOMPLETE_EQUATION)
            Return
        End If

        If s.Length = 0 Then
            StoreResult(INCOMPLETE_EQUATION)
            Return
        End If

        numberList = s.Split(sOperatorLis)

        If (numberList.Length = 1) Then
            StoreResult(numberList(0))
            Return
        End If

        For i = 0 To numberList.Length - 1
            If IsNumeric(numberList(i)) = False Then
                StoreResult(OPERAND_IS_NOT_A_NUMBER)
                Return
            End If
        Next


        If numberList.Length > 0 Then
            For i = 0 To numberList.Length - 2
                nPos += numberList(i).Length
                c = s.Substring(nPos, 1)
                operatorList.Add(c)
                nPos += 1
            Next
        End If

        If numberList.Length <> operatorList.Count + 1 Then
            StoreResult(INCOMPLETE_EQUATION)
            Return
        End If
        prevRes = numberList(0)

        For i = 0 To operatorList.Count - 1
            left = prevRes
            right = numberList(i + 1)
            oper = operatorList(i)

            prevRes = calculate_(left, right, oper)
            If IsNumeric(prevRes) = False Then
                Exit For
            End If
        Next i

        StoreResult(prevRes)

    End Sub

End Class
