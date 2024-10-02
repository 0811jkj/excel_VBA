Sub 번개런연합런()
    Dim selectedCell As Range
    
    ' Check if a single cell is selected
    If TypeName(Selection) = "Range" Then
        Set selectedCell = Selection
        ' Change the background color to gray
        selectedCell.Interior.Color = RGB(192, 192, 192) ' Gray color
        
        ' Increment the value in the selected cell by 1 (assuming it contains a number)
        If IsNumeric(selectedCell.Value) Then
            selectedCell.Value = selectedCell.Value + 1
        End If
    Else
        MsgBox "한개의 셀만 선택 되었는지 확인해주세요.", vbExclamation, "확인"
    End If
End Sub
