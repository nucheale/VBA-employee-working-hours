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
    With sheets("Осмотры")
        Set fullNameTitle = .Range(.Cells(1,1), .Cells(1, 100)).Find("ФИО")
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
        Dim fullNames as Variant
        fullNames = .Range(.Cells(fullNameTitle.Row + 1, fullNameTitle.Column), .Cells(lastRow, fullNameTitle.Column))
        fullNames = removeDublicatesFromOneDimArr(twoDimArrayToOneDim(fullNames))
        Set reportWs = Sheets.Add(After:=Sheets.Count)
        currTime = Array(Hour(Now), Minute(Now), Second(Now))
        reportWs.Name = Date & "_" & currTime(0) & "_" & currTime(1) & "_" & currTime(2)
    End With
End Sub