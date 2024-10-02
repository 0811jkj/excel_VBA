Sub 주차합산()
    Dim result As Double
    Dim lastRow As Long
    
    ' Find the last used row in column B
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 9 Then
        lastRow = 9
    End If
    
    ' Calculate the sum of values from B10 to the last non-empty cell in column B
    result = Application.WorksheetFunction.Sum(Range("B10:B" & lastRow))
    
    ' Subtract the calculated sum from the value in cell A3
    result = Range("A5").Value - result
    
    ' Write the result to cell B3
    Range("B3").Value = result
    
    ' Write the result to the next available cell in column B (starting from B10)
    Dim nextRow As Long
    nextRow = lastRow + 1
    Cells(nextRow, "B").Value = result
    
    ' Add the week label in column A corresponding to the result in column B
    Cells(nextRow, "A").Value = "" & nextRow - 10 & "주차"
    
    ' Display a message box
    MsgBox "계산 및 기록 완료", vbInformation, "확인"
End Sub
