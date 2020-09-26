Sub fizzBuzz()
    Dim num As Integer

    For i = 2 To 100
    
        num = Cells(i, 1)
        If (num Mod 5 = 0) And (num Mod 3 = 0) Then
            Cells(i, 2).Value = "Fizzbuzz"
        ElseIf num Mod 5 = 0 Then
            Cells(i, 2).Value = "Buzz"
        ElseIf num Mod 3 = 0 Then
            Cells(i, 2).Value = "Fizz"
        Else
            Cells(i, 2).Value = "."
        End If
            
    Next i


End Sub