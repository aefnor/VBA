Sub Test()
    On Error Resume Next
    
    Dim rngA As Range, cellA As Range
    Dim rngB As Range, cellB As Range
    
    'Set rng = Range("A4:A1284")
    Set rngA = Range("A1:A3")
    Set rngB = Range("B1:B3")
    Dim Year As Integer
    
    Year = 1996
    
    For Each cellA In rngA
    Dim Text As String
    Text = Right(cellA, 4)
        If (Text = Year) And (cellA.Offset(0, 1) = 1) Then
                MsgBox (cellA.Offset(0, 2))
                MsgBox (cellA)
                Exit Sub
        Else
        
        MsgBox "Not Found"
        Exit Sub
        
        End If
    Next cellA
End Sub
