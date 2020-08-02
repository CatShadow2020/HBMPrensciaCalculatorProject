Imports System.Text
Imports HBMPrensciaCalculatorProject
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub TestMethod1()
        Dim s As String
        Dim myCalc = New Calculator("0", "1", "+")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, "1")

        myCalc.SetLeft("9")
        myCalc.SetRight("1")
        myCalc.SetOperator("+")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, "10")

        myCalc.SetLeft("8")
        myCalc.SetRight("8")
        myCalc.SetOperator("*")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, "64")

        myCalc.SetLeft("1")
        myCalc.SetRight("7")
        myCalc.SetOperator("-")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, "-6")

        myCalc.SetLeft("12")
        myCalc.SetRight("3")
        myCalc.SetOperator("/")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, "4")

        myCalc.SetLeft("256")
        myCalc.SetRight("0")
        myCalc.SetOperator("/")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, ERROR_COLON + ZERO_DIVIDE)

        myCalc.SetLeft("")
        myCalc.SetRight("423")
        myCalc.SetOperator("+")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, ERROR_COLON + LEFT_PARAMETERS_IS_INVALID)

        myCalc.SetLeft("foo")
        myCalc.SetRight("423")
        myCalc.SetOperator("+")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, ERROR_COLON + LEFT_PARAMETERS_IS_INVALID)

        myCalc.SetLeft("1")
        myCalc.SetRight(Nothing)
        myCalc.SetOperator("+")
        myCalc.calculate()
        s = myCalc.getResult()
        Assert.AreEqual(s, ERROR_COLON + RIGHT_PARAMETERS_IS_INVALID)

        myCalc.SetLeft("1.7976931348623157E+308")
        myCalc.SetRight("100")
        myCalc.SetOperator("*")
        myCalc.calculate()
        s = myCalc.getResult()

        Assert.IsTrue(s.IndexOf(ERROR_COLON) = 0)


    End Sub

    <TestMethod()> Public Sub TestMethod2()
        Dim s As String
        Dim myCalc2 As CalculatorPart2 = New CalculatorPart2()

        myCalc2.calculate("1+1+1+1+1+1+1+!+1+1")
        s = myCalc2.getResult()
        Assert.AreEqual(s, ERROR_COLON + OPERAND_IS_NOT_A_NUMBER)

        myCalc2.calculate("1*2+4-500")
        s = myCalc2.getResult()
        Assert.AreEqual(s, "-494")

        myCalc2.calculate("")
        s = myCalc2.getResult()
        Assert.AreEqual(s, ERROR_COLON + INCOMPLETE_EQUATION)

        myCalc2.calculate(Nothing)
        s = myCalc2.getResult()
        Assert.AreEqual(s, ERROR_COLON + INCOMPLETE_EQUATION)

        myCalc2.calculate("1+")
        s = myCalc2.getResult()
        Assert.AreEqual(s, ERROR_COLON + OPERAND_IS_NOT_A_NUMBER)

        myCalc2.calculate("100*1.7976931348623157E+308")
        s = myCalc2.getResult()
        Assert.IsTrue(s.IndexOf(ERROR_COLON) = 0)

        myCalc2.calculate("1/0")
        s = myCalc2.getResult()
        Assert.AreEqual(s, ERROR_COLON + ZERO_DIVIDE)

        For i = 0 To 11
            myCalc2.calculate("1+" + i)
        Next i


        s = myCalc2.getResult()
        Assert.AreEqual(s, "12")

        For i = 1 To 10
            s = myCalc2.getPreviousResult(i)
            Assert.AreEqual(s, Convert.ToString(12 - i))
        Next i


    End Sub

End Class