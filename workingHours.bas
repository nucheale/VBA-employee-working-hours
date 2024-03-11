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

Sub Расчет()
    maxWorkingHours = 16
    maxWorkingMinutes = maxWorkingHours * 60

    With Sheets("Осмотры")
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
        reportWs.Cells(1, 1) = "Список водителей с количеством отработанных дней и часов"
        reportWs.Cells(1, 2) = "Отработано дней"
        reportWs.Cells(1, 3) = "Отработано часов"
        reportWs.Cells(2, 1).Resize(UBound(fullNames), 1).Value = Application.Transpose(fullNames)
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
            fullResult(workerId, 1) = wName
            For i = LBound(fullData, 1) To UBound(fullData, 1)
            'надо все переделать. если предрейсовый повторяется 2 раза - сразу +12ч и переходить к следующей дате как к первой

                If date1Found = False Then
                    ' If wName = fullData(i, fullNameTitle.Column) And ((LCase(fullData(i, 11)) = "допущен") Or (LCase(fullData(i, 11)) = "прошёл")) Then
                    If wName = fullData(i, fullNameTitle.Column) And LCase(fullData(i, 11)) = "допущен" Then
                        date1 = fullData(i, 1)
                        date1Found = True
                        date1Counter = date1Counter + 1
                        startWork = fullData(i, 2)
                        data1Index = i
                    End If
                ElseIf date1Found = True Then
                    If wName = fullData(i, fullNameTitle.Column) And LCase(fullData(i, 11)) = "прошёл" Then
                        If (LCase(fullData(data1Index, 6)) = "предрейсовый") And (LCase(fullData(i, 6)) = LCase(fullData(data1Index, 6))) Then
                            Debug.Print "gfdg"
                            date1 = fullData(i, 1)
                            date2 = Empty
                            date1Found = True
                            date2Found = False
                            GoTo nextIteration
                        Else
                            date2 = fullData(i, 1)
                            'date1Found = False
                            date2Found = True
                            date2Counter = date2Counter + 1
                            endWork = fullData(i, 2)
                        End If
                    End If
                End If
                If date1Found And date2Found Then
                    workingMinutes = DateDiff("n", startWork, endWork)
                    workedHours = workingMinutes / 60
                    standardWorkingHours = 12
                    If workingMinutes <= maxWorkingMinutes Then
                        ' Debug.Print "startWork: ", startWork
                        ' Debug.Print "endWork: ", endWork
                        ' Debug.Print workedHours
                        fullResult(workerId, 2) = fullResult(workerId, 2) + workedHours
                        If fullResult(workerId, 1) Like "*Мулев*" Then Debug.Print fullResult(workerId, 1), ": +", workedHours
                        date1Found = False
                        date2Found = False
                    Else
                        fullResult(workerId, 2) = fullResult(workerId, 2) + standardWorkingHours
                        If fullResult(workerId, 1) Like "*Мулев*" Then Debug.Print fullResult(workerId, 1), ": +", standardWorkingHours
                        ' date1 = date2
                        date1 = Empty
                        date2 = Empty
                        date1Found = False
                        date2Found = False
                        GoTo nextIteration
                    End If
                End If
nextIteration:
            Next i
            ' Debug.Print wName
            ' Debug.Print date1Counter
            ' Debug.Print date2Counter
        workerId = workerId + 1
        Next wName
    End With
    
    With reportWs
        .Cells(2, 1).Resize(UBound(fullResult), UBound(fullResult, 2)).Value = fullResult
    End With
End Sub

