Sub QuadraticSolver()
    Dim a As Double, b As Double, c As Double
    Dim to_root As Double
    Dim x1 As Double, x2 As Double
    Dim result As String
    Dim cellRange As Range
    
    Set cellRange = Range("C1:C6")
    
    cellRange.NumberFormat = "0.###################"
    
    Range("A1").Value = ""
    Range("A2").Value = ""
    Range("A3").Value = ""
    Range("A5").Value = ""
    Range("A6").Value = ""
    
    Range("C1").Value = ""
    Range("C2").Value = ""
    Range("C3").Value = ""
    Range("C5").Value = ""
    Range("C6").Value = ""

    a = InputBox("Enter coefficient a:", "Quadratic Solver")
    b = InputBox("Enter coefficient b:", "Quadratic Solver")
    c = InputBox("Enter coefficient c:", "Quadratic Solver")
    
    Range("A1").Value = "Coefficient A:"
    Range("A2").Value = "Coefficient B:"
    Range("A3").Value = "Coefficient C:"
    
    Range("C1").Value = a
    Range("C2").Value = b
    Range("C3").Value = c
    
    to_root = b ^ 2 - 4 * a * c
    
    If to_root > 0 Then
        x1 = (-b + Sqr(to_root)) / (2 * a)
        x2 = (-b - Sqr(to_root)) / (2 * a)
        result = "The solutions are: x1 = " & x1 & " and x2 = " & x2
        Range("A5").Value = "Solution 1:"
        Range("A6").Value = "Solution 2:"
        Range("C5").Value = x1
        Range("C6").Value = x2
    ElseIf to_root = 0 Then
        x1 = -b / (2 * a)
        result = "The solution is: x = " & x1
        Range("A5").Value = "Solution:"
        Range("C5").Value = x1
    Else
        result = "The equation has no real solutions."
        Range("A5").Value = "No real solution found."
    End If
    
    MsgBox result, vbInformation, "Quadratic Solver"
    Columns("C:C").AutoFit
End Sub
