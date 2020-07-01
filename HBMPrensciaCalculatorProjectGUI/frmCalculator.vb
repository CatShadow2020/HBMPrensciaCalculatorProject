Imports HBMPrensciaCalculatorProject

Public Class frmCalculator
    Dim m_Calculator As CalculatorPart2
    Dim m_Insert As Boolean

    '
    ' Insert into display operator
    '
    '

    Private Sub InsertOperatorIntoDisplay(c As String)
        Dim txt As String = DisplayTextBox.Text
        If txt.Length = 0 Or m_Insert = False Then
            Return
        End If
        Dim op As String
        op = txt(txt.Length - 1)
        If op = "*" Or op = "/" Or op = "+" Or op = "-" Then
            txt = txt.Remove(txt.Length - 1)
            txt += c
            DisplayTextBox.Text = txt
            Return
        End If

        InsertIntoDisplay(c)

    End Sub
    '
    ' Insert into display digit
    '
    '
    Private Sub InsertIntoDisplay(c As String)
        If m_Insert Then
            Dim txt As String = DisplayTextBox.Text
            txt = txt + c
            DisplayTextBox.Text = txt
        Else
            DisplayTextBox.Text = c
            m_Insert = True
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        InsertIntoDisplay("7")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        InsertIntoDisplay("8")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        InsertIntoDisplay("9")
    End Sub

    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        DisplayTextBox.Text = ""
    End Sub

    Private Sub ButtonBackSpace_Click(sender As Object, e As EventArgs) Handles ButtonBackSpace.Click
        Dim txt As String = DisplayTextBox.Text
        If m_Insert = False Then
            DisplayTextBox.Text = ""
            Return
        End If
        If txt.Length > 0 Then
            txt = txt.Substring(0, txt.Length - 1)
            DisplayTextBox.Text = txt
        End If
    End Sub

    Private Sub ButtonMultiply_Click(sender As Object, e As EventArgs) Handles ButtonMultiply.Click
        InsertOperatorIntoDisplay("*")
    End Sub


    Private Sub ButtonDivide_Click(sender As Object, e As EventArgs) Handles ButtonDivide.Click
        InsertOperatorIntoDisplay("/")
    End Sub

    Private Sub ButtonAdd_Click(sender As Object, e As EventArgs) Handles ButtonAdd.Click
        InsertOperatorIntoDisplay("+")
    End Sub

    Private Sub ButtonSubtract_Click(sender As Object, e As EventArgs) Handles ButtonSubtract.Click
        InsertOperatorIntoDisplay("-")
    End Sub

    Private Sub Button0_Click(sender As Object, e As EventArgs) Handles Button0.Click
        InsertIntoDisplay("0")
    End Sub

    Private Sub ButtonEnter_Click(sender As Object, e As EventArgs) Handles ButtonEnter.Click
        Dim txt As String = DisplayTextBox.Text
        If txt.Length > 0 Then
            m_Calculator.calculate(txt)
            txt = m_Calculator.getResult()
            DisplayTextBox.Text = txt
            m_Insert = False
            UpdateHistoryListBox()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        InsertIntoDisplay("4")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        InsertIntoDisplay("5")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        InsertIntoDisplay("6")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        InsertIntoDisplay("2")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        InsertIntoDisplay("3")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        InsertIntoDisplay("1")
    End Sub

    Private Sub frmCalculator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        m_Calculator = New CalculatorPart2
    End Sub

    Private Sub UpdateHistoryListBox()
        Dim txt As String
        Dim i As Integer

        Me.HistoryListBox.Items.Clear()

        txt = m_Calculator.getResult()

        If IsNothing(txt) Then
            Return
        End If

        Me.HistoryListBox.Items.Add(txt)

        For i = 1 To 10

            txt = m_Calculator.getPreviousResult(i)
            If IsNothing(txt) Then
                Return
            End If


            Me.HistoryListBox.Items.Add(txt)
        Next i
    End Sub

End Class
