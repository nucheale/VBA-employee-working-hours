Function twoDimArrayToOneDim(oldArr)
    Dim newArr As Variant
    ReDim newArr(1 To UBound(oldArr, 1) * UBound(oldArr, 2))
    For i = LBound(oldArr, 1) To UBound(oldArr, 1)
        newArr(i) = oldArr(i, 1)
    Next i
    twoDimArrayToOneDim = newArr
End Function

Function removeDublicatesFromOneDimArr(arr)
    Dim coll As New Collection
    For Each e In arr
        On Error Resume Next
        coll.Add e, e
        On Error GoTo 0
    Next e
    Dim uniqueArr As Variant
    ReDim uniqueArr(1 To coll.Count)
    For i = 1 To coll.Count
        uniqueArr(i) = coll(i)
    Next i
    removeDublicatesFromOneDimArr = uniqueArr
End Function

Function convertToTimeFormat(hours) As String
    hh = Int(hours)
    mm = Round((hours - hh) * 60)
    convertToTimeFormat = Format(hh, "00") & ":" & Format(mm, "00")
End Function

Sub Расчет()
    maxWorkingHours = 16
    maxWorkingMinutes = maxWorkingHours * 60
    standardWorkingHours = 12

    Set dataWs = Sheets("Осмотры")
    With dataWs
        If dataWs.AutoFilterMode Then dataWs.AutoFilterMode = False
        'If Worksheets("Осмотры").ShowAllData = False Then Worksheets("Осмотры").ShowAllData
        Set fullNameTitle = .Range(.Cells(1, 1), .Cells(1, 100)).Find("ФИО")
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
        lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
        Dim fullNames As Variant
        fullNames = .Range(.Cells(fullNameTitle.Row + 1, fullNameTitle.Column), .Cells(lastRow, fullNameTitle.Column))
        fullNames = removeDublicatesFromOneDimArr(twoDimArrayToOneDim(fullNames))
        Dim fullData As Variant
        fullData = .Range(.Cells(fullNameTitle.Row + 1, 1), .Cells(lastRow, lastColumn))

        ' --------------- новый лист с отчетом ---------------
        Set reportWs = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        currTime = Array(Hour(Now), Minute(Now), Second(Now))
        reportWs.Name = Date & "_" & currTime(0) & "_" & currTime(1) & "_" & currTime(2)
        reportWs.Cells(1, 1) = "ФИО"
        reportWs.Cells(1, 2) = "Отработано дней"
        reportWs.Cells(1, 3) = "Отработано часов"
        ' reportWs.Cells(2, 1).Resize(UBound(fullNames), 1).Value = Application.Transpose(fullNames)
        ' --------------- конец новый лист с отчетом ---------------
        
        Dim fullResult() As Variant
        ReDim fullResult(1 To UBound(fullNames), 1 To 3)
        workerId = 1
        For Each wName In fullNames
            date1Counter = 0
            date2Counter = 0
            date1Found = False
            date2Found = False
            date1 = Empty
            date2 = Empty
            onlyOneDate = False
            onlyOneDate1 = False
            onlyOneDate2 = False
            fullResult(workerId, 1) = wName
            For i = LBound(fullData, 1) To UBound(fullData, 1)
                If wName = fullData(i, fullNameTitle.Column) Then
                    If LCase(fullData(i, 6)) = "предрейсовый" And LCase(fullData(i, 11)) = "допущен" Then
                        If date1Found = False Then
                            date1 = fullData(i, 2)
                            date1Found = True
                        Else
                            workedHours = standardWorkingHours
                            onlyOneDate = True
                            onlyOneDate1 = True
                            GoTo writingResult
                        End If
                    ElseIf LCase(fullData(i, 6)) = "послерейсовый" And LCase(fullData(i, 11)) = "прошёл" Then
                        If date1Found And Not date2Found Then
                            date2 = fullData(i, 2)
                            date2Found = True
                        ElseIf Not date1Found And Not date2Found Then
                            workedHours = standardWorkingHours
                            onlyOneDate = True
                            onlyOneDate2 = True
                            GoTo writingResult
                        End If
                    End If
                End If
              
writingResult:
                If (date1Found And date2Found) Or onlyOneDate Then
                    If workedHours = Empty Then
                        workingMinutes = DateDiff("n", date1, date2)
                        workedHours = workingMinutes / 60
                    End If
                    ' Debug.Print "date1: ", date1
                    ' Debug.Print "date2: ", date2
                    ' Debug.Print workedHours
                    If workedHours <= maxWorkingHours Then workedHours = standardWorkingHours
                    fullResult(workerId, 2) = fullResult(workerId, 2) + 1
                    fullResult(workerId, 3) = fullResult(workerId, 3) + workedHours
                    date1 = Empty
                    date2 = Empty
                    If Not onlyOneDate1 Then date1Found = False
                    If Not onlyOneDate2 Then date2Found = False
                    workedHours = Empty
                    If onlyOneDate1 Then date1 = fullData(i, 2)
                    onlyOneDate = False
                    onlyOneDate1 = False
                    onlyOneDate2 = False
                End If
            Next i
            workerId = workerId + 1
        Next wName
    End With
    
    ' For i = LBound(fullResult) To UBound(fullResult)
    '     fullResult(i, 3) = convertToTimeFormat(fullResult(i, 3))
    ' Next i

    With reportWs
        .Cells(2, 1).Resize(UBound(fullResult), UBound(fullResult, 2)).Value = fullResult
        .Cells.EntireColumn.AutoFit
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
        lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(lastRow, lastColumn)).Borders.LineStyle = xlContinuous
        .Range(.Cells(2, lastColumn), .Cells(lastRow, lastColumn)).NumberFormat = "#,##0.00"
        .Range(.Cells(1, 1), .Cells(1, lastColumn)).Font.Bold = True
    End With
End Sub