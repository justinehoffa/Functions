Option Strict On
Option Explicit On

'Justine Hoffa
'RCET0265
'Fall2020
'Functions
'https://github.com/justinehoffa/Functions

Module Functions

    Sub Main()
        Dim aLetter As String
        aLetter = "Q"
        'Console.WriteLine(aLetter)
        'Console.ReadLine()

        'Console.WriteLine(Letter())
        'Console.ReadLine()

        'Console.WriteLine(DoesNotReturnStuff())
        'Console.ReadLine()

        'Console.WriteLine(AddNumbers(3I, 7I))
        'Console.ReadLine()

        'For i = 1 To 3
        '    Console.WriteLine(RunningTotal(5))
        '    Console.ReadLine()
        'Next

        AccumulateMessage("Hello Johnny!")
        AccumulateMessage("That's Not a Number")
        AccumulateMessage("USer Name is required")
        MsgBox(AccumulateMessage(""))
    End Sub

    Function Letter() As String
        Dim someOtherLetter As String
        someOtherLetter = "Y"
        Return someOtherLetter
    End Function

    Sub DoesNotReturnStuff()
        Dim someOtherLetter As String
        someOtherLetter = "X"
        'Return someOtherLetter
    End Sub

    Function AddNumbers(ByVal firstnumber As Integer, ByVal secondNumber As Integer) As Integer
        Return firstnumber + secondNumber
    End Function

    Function RunningTotal(ByVal currentAmount As Decimal) As Decimal
        Static total As Decimal
        total += currentAmount
        Return total
    End Function

    Function AccumulateMessage(ByVal newMessage As String) As String
        Static userMessage As String
        userMessage &= newMessage & vbNewLine
        Return userMessage
    End Function
End Module
