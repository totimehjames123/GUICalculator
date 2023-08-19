Imports System.Numerics
Imports Microsoft.VisualBasic.FileIO

Public Class Form1
    Private currentCalculation As String = ""

    Private Sub NumberButton_Click(sender As Object, e As EventArgs) Handles zeroButton.Click, oneButton.Click, twoButton.Click, threeButton.Click, fourButton.Click, fiveButton.Click, sixButton.Click, sevenButton.Click, eightButton.Click, nineButton.Click
        Dim clickedButton As Button = CType(sender, Button)
        currentCalculation += clickedButton.Text
        inputField.Text = currentCalculation
    End Sub

    Private Sub OperatorButton_Click(sender As Object, e As EventArgs) Handles plusButton.Click, minusButton.Click, timesButton.Click, divisionButton.Click
        Dim clickedButton As Button = CType(sender, Button)

        If clickedButton.Text = "X" Then
            currentCalculation += " " + "*" + " "
            inputField.Text = currentCalculation
        Else
            currentCalculation += " " + clickedButton.Text + " "
            inputField.Text = currentCalculation

        End If

    End Sub

    Private Sub dotButton_Click(sender As Object, e As EventArgs) Handles dotButton.Click
        If Not currentCalculation.EndsWith(".") Then
            currentCalculation += "."
            inputField.Text = currentCalculation
        End If
    End Sub

    Private Sub openBracketButton_Click(sender As Object, e As EventArgs) Handles openBracketButton.Click
        currentCalculation += "("
        inputField.Text = currentCalculation
    End Sub

    Private Sub closeBracketButton_Click(sender As Object, e As EventArgs) Handles closeBracketButton.Click
        currentCalculation += ")"
        inputField.Text = currentCalculation
    End Sub

    Private Sub equaltoButton_Click(sender As Object, e As EventArgs) Handles equaltoButton.Click
        Try
            If IsValidExpression(currentCalculation) Then
                Dim result As Double = EvaluateExpression(currentCalculation)
                inputField.Text = result.ToString()
                currentCalculation = result.ToString()
            End If
        Catch ex As Exception
            inputField.Text = "Syntax error"
        End Try


    End Sub



    Private Function EvaluateExpression(expression As String) As Double
        Try
            If IsValidExpression(currentCalculation) Then
                Dim dt As New DataTable()
                Dim result As Object = dt.Compute(expression, currentCalculation)
                Return Convert.ToDouble(result)
            End If
        Catch ex As Exception
            inputField.Text = "Syntax error!"
        End Try


    End Function

    Private Sub clearButton_Click_1(sender As Object, e As EventArgs) Handles clearButton.Click
        inputField.Clear()
        currentCalculation = ""
    End Sub

    Private Function CalculateFactorial(number As BigInteger) As BigInteger
        Try
            Dim result As BigInteger = 1

            For i As BigInteger = 2 To number
                result *= i
            Next

            Return result
        Catch ex As Exception
            inputField.Text = "Stack error!"
        End Try

    End Function


    Private Async Sub factorialButton_Click(sender As Object, e As EventArgs) Handles factorialButton.Click
        Try
            If BigInteger.TryParse(currentCalculation, Nothing) Then
                Dim number As BigInteger = BigInteger.Parse(currentCalculation)

                If number < 0 Then
                    MessageBox.Show("Factorial is not defined for negative numbers.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim factorialResult As BigInteger = Await Task.Run(Function() CalculateFactorial(number))
                inputField.Text = factorialResult.ToString()
                currentCalculation = factorialResult.ToString()
            Else
                MessageBox.Show("Invalid input for factorial calculation.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Function IsValidExpression(expression As String) As Boolean


        If String.IsNullOrWhiteSpace(expression) Then
            Return False
        End If

        Dim lastChar As Char = expression(expression.Length - 1)
        If "+-*/.".Contains(lastChar) Then
            Return False
        End If

        Dim openParenCount As Integer = expression.Count(Function(c) c = "(")
        Dim closeParenCount As Integer = expression.Count(Function(c) c = ")")

        If openParenCount <> closeParenCount Then
            Return False
        End If

        Return True
    End Function

End Class
