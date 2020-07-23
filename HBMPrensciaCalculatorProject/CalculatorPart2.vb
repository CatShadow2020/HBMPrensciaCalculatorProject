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

    Private Function OperatorSignHelper(ByRef operatorList As List(Of String), ByRef numberList() As String) As Boolean
        Dim aNumberList As ArrayList = New ArrayList(numberList)
        Dim aOperatorList As ArrayList = New ArrayList(operatorList)
        Dim left, right, oper As String
        Dim res As String
        Dim i As Integer

        i = 0


        While i < aOperatorList.Count
            left = aNumberList(i)
            right = aNumberList(i + 1)
            oper = aOperatorList(i)

            If oper = "-" And left = "" Then
                aNumberList(i) = "-" + aNumberList(i + 1)
                aOperatorList.RemoveAt(i)
                aNumberList.RemoveAt(i + 1)
                Else
                    i += 1
            End If


        End While

        operatorList.Clear()

        For i = 0 To aOperatorList.Count - 1
            operatorList.Add(aOperatorList(i))
        Next i

        ReDim numberList(aNumberList.Count - 1)
        For i = 0 To aNumberList.Count - 1
            numberList(i) = aNumberList(i)
        Next i
        Return True

    End Function

    Private Function OperatorPrecedenceHelper(ByRef operatorList As List(Of String), ByRef numberList() As String) As Boolean
        Dim aNumberList As ArrayList = New ArrayList(numberList)
        Dim aOperatorList As ArrayList = New ArrayList(operatorList)
        Dim left, right, oper As String
        Dim res As String
        Dim i As Integer

        i = 0


        While i < aOperatorList.Count
            left = aNumberList(i)
            right = aNumberList(i + 1)
            oper = aOperatorList(i)

            If oper = "*" Or oper = "/" Then
                res = calculate_(left, right, oper)
                If IsNumeric(res) = False Then
                    StoreResult(res)
                    Return False
                End If

                aNumberList(i) = res
                aOperatorList.RemoveAt(i)
                aNumberList.RemoveAt(i + 1)
            Else
                i += 1
            End If


        End While

        operatorList.Clear()

        For i = 0 To aOperatorList.Count - 1
            operatorList.Add(aOperatorList(i))
        Next i

        ReDim numberList(aNumberList.Count - 1)
        For i = 0 To aNumberList.Count - 1
            numberList(i) = aNumberList(i)
        Next i
        Return True

    End Function


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

        If numberList.Length > 0 Then
            For i = 0 To numberList.Length - 2
                nPos += numberList(i).Length
                c = s.Substring(nPos, 1)
                operatorList.Add(c)
                nPos += 1
            Next
        End If

        OperatorSignHelper(operatorList, numberList)

        For i = 0 To numberList.Length - 1
            If IsNumeric(numberList(i)) = False Then
                StoreResult(OPERAND_IS_NOT_A_NUMBER)
                Return
            End If
        Next


        If numberList.Length <> operatorList.Count + 1 Then
            StoreResult(INCOMPLETE_EQUATION)
            Return
        End If

        If OperatorPrecedenceHelper(operatorList, numberList) = False Then Return

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
