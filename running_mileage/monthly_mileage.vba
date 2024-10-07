Sub 월별마일리지()
    Dim lastRow As Long
    Dim currentRow As Long
    Dim monthCount As Long
    Dim mileage As Double
    Dim resultArray(1 To 50) As Double
    Dim nameArray(1 To 50) As String
    Dim topMileage(1 To 5) As Double
    Dim topNames(1 To 5) As String
    Dim i As Long
    Dim previousSum As Double
    
    ' 마지막 행 찾기
    lastRow = Cells(Rows.Count, "BE").End(xlUp).Row
    
    ' 첫 실행 시 4행부터 시작, 그 후에는 다음 행으로 이동
    If lastRow < 4 Then
        currentRow = 4 ' 처음 실행 시는 4행
    Else
        currentRow = lastRow + 1 ' 그 다음 실행 시는 마지막 행 아래
    End If
    
    ' 달 수 계산
    monthCount = currentRow - 3
    Cells(currentRow, "BD").Value = monthCount & "달차"
    
   ' D열부터 BA열까지 계산하여 BE부터 DB까지 값 작성
    For i = 4 To 53 ' D = 4, BA = 53
        ' 이전 달차의 합계를 계산
        If currentRow > 4 Then
            previousSum = Application.WorksheetFunction.Sum(Range(Cells(4, i + 53), Cells(currentRow - 1, i + 53))) ' 이전 달차 합계를 가져옴
        Else
            previousSum = 0 ' 첫 번째 달차일 경우
        End If
        
        mileage = Cells(3, i).Value - previousSum ' D3부터 BA3까지의 누적 거리에서 이전 달차의 합계를 뺌
        Cells(currentRow, i + 53).Value = mileage ' BE~DB까지 마일리지 작성
        resultArray(i - 3) = mileage ' 계산된 마일리지를 배열에 저장
        nameArray(i - 3) = Cells(2, i).Value ' 이름을 배열에 저장
    Next i
    
    ' Top 5 초기화
    For i = 1 To 5
        topMileage(i) = -1
        topNames(i) = ""
    Next i
    
    ' Top 5 선정
    For i = 1 To 50
        Call SortTop5(resultArray(i), nameArray(i), topMileage, topNames)
    Next i
    
    ' Top 5 결과 메시지 박스 출력
    Dim msg As String
    msg = "월별 개인 마일리지 Top 5:" & vbCrLf
    For i = 1 To 5
        msg = msg & i & ". " & topNames(i) & " - " & topMileage(i) & " km" & vbCrLf
    Next i
    MsgBox msg, vbInformation, "Top 5 개인 마일리지"
End Sub

' Top 5 정렬 함수
Sub SortTop5(newMileage As Double, newName As String, ByRef topMileage() As Double, ByRef topNames() As String)
    Dim i As Long
    For i = 1 To 5
        If newMileage > topMileage(i) Then
            ' 순위 낮은 것들 아래로 이동
            If i < 5 Then topMileage(5) = topMileage(4): topNames(5) = topNames(4)
            If i < 4 Then topMileage(4) = topMileage(3): topNames(4) = topNames(3)
            If i < 3 Then topMileage(3) = topMileage(2): topNames(3) = topNames(2)
            If i < 2 Then topMileage(2) = topMileage(1): topNames(2) = topNames(1)
            ' 새로운 상위 마일리지 및 이름 삽입
            topMileage(i) = newMileage
            topNames(i) = newName
            Exit For
        End If
    Next i
End Sub

