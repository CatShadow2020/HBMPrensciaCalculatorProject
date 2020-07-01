Module Module1

    Sub Main()
        Dim myCalc = New Calculator("0", "1", "+")
        myCalc.calculate()
        myCalc.setLeft("1")
        myCalc.calculate()
        myCalc.setLeft("2")
        myCalc.calculate()
        myCalc.setLeft("3")
        myCalc.calculate()
        myCalc.setLeft("4")
        myCalc.calculate()
        myCalc.setLeft("5")
        myCalc.calculate()
        myCalc.setLeft("6")
        myCalc.calculate()
        myCalc.setLeft("7")
        myCalc.calculate()
        myCalc.setLeft("8")
        myCalc.calculate()
        myCalc.setLeft("9")
        myCalc.calculate()
        myCalc.setLeft("10")
        myCalc.calculate()
        myCalc.setLeft("11")
        myCalc.calculate()
        myCalc.setLeft("12")
        myCalc.calculate()
        Console.WriteLine("5 + 1 = " + myCalc.getResult())

        myCalc.SetLeft(Nothing)
        myCalc.calculate()
        Console.WriteLine("Result:" + myCalc.getResult())

        Dim myCalc2 As CalculatorPart2 = New CalculatorPart2()

        myCalc2.calculate("1*2+4-500")
        Console.WriteLine("Result:" + myCalc2.getResult())

    End Sub

End Module
