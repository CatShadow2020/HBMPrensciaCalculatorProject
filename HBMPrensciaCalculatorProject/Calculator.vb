'1.1 Part 1: Basic Implementation 
' 
'The Class must implement the following methods: 
' 
'  Calculator(left, right, operator) 
'    - Constructor, all three parameters are strings 
'    - valid values for left And right are numbers 
'    - valid values for operator are: "+", "-", "*", "/" 



Public Class Calculator
    Inherits ResultClass
    Private m_Left As String
    Private m_Right As String
    Private m_Operator As String



    Public Sub New(left As String, right As String, oper As String)
        MyBase.New(10)
        m_Left = left
        m_Right = right
        m_Operator = oper
    End Sub


    Public Sub SetLeft(value As String)
        m_Left = value
    End Sub

    Public Sub SetRight(value As String)
        m_Right = value
    End Sub


    Public Sub SetOperator(value As String)
        m_Operator = value
    End Sub

    Public Sub calculate()
        Dim resultString As String
        resultString = calculate_(m_Left, m_Right, m_Operator)
        StoreResult(resultString)
    End Sub
End Class
