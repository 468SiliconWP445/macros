Sub QuadraticSolver()
    Dim a As Double, b As Double, c As Double
    Dim to_root As Double
    Dim x1 As Double, x2 As Double
    Dim result As String

    a = InputBox("Enter coefficient a:", "Quadratic Solver")
    b = InputBox("Enter coefficient b:", "Quadratic Solver")
    c = InputBox("Enter coefficient c:", "Quadratic Solver")
    
    Range("A1").Value = "Coefficient A:"
    Range("A2").Value = "Coefficient B:"
    Range("A3").Value = "Coefficient C:"
    
    Range("B1").Value = a
    Range("B2").Value = b
    Range("B3").Value = c
    
    to_root = b ^ 2 - 4 * a * c
    
    If to_root > 0 Then
        x1 = (-b + Sqr(to_root)) / (2 * a)
        x2 = (-b - Sqr(to_root)) / (2 * a)
        result = "The solutions are: x1 = " & x1 & " and x2 = " & x2
        Range("A5").Value = "Solution 1:"
        Range("A6").Value = "Solution 2:"
        Range("B5").Value = x1
        Range("B6").Value = x2
    ElseIf to_root = 0 Then
        x1 = -b / (2 * a)
        result = "The solution is: x = " & x1
        Range("A5").Value = "Solution:"
        Range("B5").Value = x1
    Else
        result = "The equation has no real solutions."
        Range("A5").Value = "No real solution found."
    End If
    
    MsgBox result, vbInformation, "Quadratic Solver"
End Sub
