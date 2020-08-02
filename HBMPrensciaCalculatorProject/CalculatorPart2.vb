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

    Public Sub calculate(s As String)
        Dim varDictionary As New Dictionary(Of String, ParserVariable)
        Dim sResult, ss As String

        If IsNothing(s) Then
            StoreResult(ERROR_COLON + INCOMPLETE_EQUATION)
            Return
        End If

        s.Replace(" ", "")

        If (CheckParentheses(s) = False) Then
            StoreResult(ERROR_COLON + PARENTHESES_ARE_NOT_VALID)
            Return
        End If


        parser_(s, Nothing, varDictionary)

        Dim list As New List(Of String)(varDictionary.Keys)

        Do While IsNothing(varDictionary("A0").value)
            For Each iKey As String In list
                ss = varDictionary(iKey).value
                s = varDictionary(iKey).formula
                If IsNothing(ss) Then
                    sResult = calculate0(s, varDictionary)
                    If IsNothing(sResult) = False Then
                        If sResult.IndexOf(ERROR_COLON) = 0 Then
                            StoreResult(sResult)
                            Return
                        End If
                    End If
                    varDictionary(iKey).value = sResult
                End If

            Next

        Loop

        StoreResult(varDictionary("A0").value)
    End Sub



    Private Sub parser_(s As String, sCurrentVar As String,
                           ByRef varDictionary As Dictionary(Of String, ParserVariable))
        Dim i As Integer
        Dim sc As String
        Dim parenthesisCounter As Integer = 0
        Dim nLastNumber As Integer
        Dim leftString, sOperator As String
        Dim mainString As String = ""

        nLastNumber = varDictionary.Count

        If (IsNothing(sCurrentVar)) Then
            sCurrentVar = "A0"
        End If

        leftString = ""
        For i = 0 To s.Length - 1
            sc = s(i)

            If sc = "(" Then

                If (parenthesisCounter > 0) Then
                    leftString = leftString + sc
                End If

                parenthesisCounter += 1
                Continue For
            End If

            If sc = ")" Then
                parenthesisCounter -= 1

                If (parenthesisCounter > 0) Then
                    leftString = leftString + sc
                End If

                If (parenthesisCounter = 0) Then
                    nLastNumber = varDictionary.Count
                    Dim sName As String = "A" + Convert.ToString(nLastNumber + 1)

                    varDictionary(sName) = New ParserVariable(leftString, Nothing)

                    parser_(leftString, sName,
                           varDictionary)

                    leftString = sName

                    Continue For
                End If

                Continue For
            End If

            If (parenthesisCounter > 0) Then
                leftString = leftString + sc
                Continue For
            End If

            If isOperator(sc) Then
                sOperator = sc

                mainString += leftString
                mainString += sc

                leftString = ""

                Continue For
            End If


            leftString = leftString + sc
        Next i

        If leftString.Length > 0 Then
            mainString += leftString
        End If

        varDictionary(sCurrentVar) = New ParserVariable(mainString, Nothing)
    End Sub

    Private Function isOperator(s As String) As Boolean
        If s = "*" Or s = "/" Or s = "+" Or s = "-" Then
            Return True
        End If
        Return False
    End Function


    '
    ' Calculates equation 's' :
    ' The format of an equation is 
    ' "<number> <operator> <number2> <operator2> ... <numberN> <operatorN>", 
    ' starting And ending with a digit. 
    '
    '




    Private Function calculate0(s As String,
                               ByRef varDictionary As Dictionary(Of String, ParserVariable)
                               ) As String
        Dim sOperatorLis() As Char = {"*", "/", "+", "-"}


        Dim operatorList As New List(Of String)()
        Dim numberList() As String
        Dim nPos As Integer = 0
        Dim c As Char
        Dim i As Integer
        Dim left, right, oper As String
        Dim prevRes As String

        If IsNothing(s) Then
            'StoreResult(INCOMPLETE_EQUATION)
            Return ERROR_COLON + INCOMPLETE_EQUATION
        End If

        If s.Length = 0 Then
            'StoreResult(INCOMPLETE_EQUATION)
            Return ERROR_COLON + INCOMPLETE_EQUATION
        End If

        numberList = s.Split(sOperatorLis)

        If (numberList.Length = 1) Then
            'StoreResult(numberList(0))
            If varDictionary.ContainsKey(numberList(0)) = True Then
                Return varDictionary(numberList(0)).value
            End If
            Return numberList(0)
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
                If varDictionary.ContainsKey(numberList(i)) = False Then
                    'StoreResult(OPERAND_IS_NOT_A_NUMBER)
                    Return ERROR_COLON + OPERAND_IS_NOT_A_NUMBER
                End If
            End If
        Next


        For i = 0 To numberList.Length - 1
            If IsNumeric(numberList(i)) = False Then
                If varDictionary.ContainsKey(numberList(i)) = True Then
                    Dim sValue As String = varDictionary(numberList(i)).value
                    If IsNothing(sValue) Then
                        Return Nothing
                    Else
                        numberList(i) = sValue
                    End If
                Else
                    'StoreResult(OPERAND_IS_NOT_A_NUMBER)
                    Return ERROR_COLON + OPERAND_IS_NOT_A_NUMBER
                End If
            End If
        Next


        If numberList.Length <> operatorList.Count + 1 Then
            'StoreResult(INCOMPLETE_EQUATION)
            Return ERROR_COLON + INCOMPLETE_EQUATION
        End If

        Dim sError As String = ""
        If OperatorPrecedenceHelper(operatorList, numberList, sError) = False Then Return sError

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

        'StoreResult(prevRes)
        Return prevRes

    End Function

    Private Function OperatorSignHelper(ByRef operatorList As List(Of String), ByRef numberList() As String) As Boolean
        Dim aNumberList As ArrayList = New ArrayList(numberList)
        Dim aOperatorList As ArrayList = New ArrayList(operatorList)
        Dim left, right, oper As String
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

    Private Function OperatorPrecedenceHelper(ByRef operatorList As List(Of String), ByRef numberList() As String, ByRef sError As String) As Boolean
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
                    'StoreResult(res)
                    sError = res
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
End Class
