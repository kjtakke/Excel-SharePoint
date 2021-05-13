Private Function arrayDimentionCounter(index As Variant) As Integer
    'This Function Counts the Columns/Dimentions in an Array
    'index is the input array
    
        On Error GoTo LC:
        For L = 1 To 100
            TempVar = index(1, L)
        Next L
LC:
        L = L - 1
        On Error GoTo 0
        arrayDimentionCounter = L
End Function




Private Function uniqueValues_FromBaseOneArray(dataArray As Variant, dimention As Integer) As Variant
    Dim uniqueCollection As Collection: Set uniqueCollection = New Collection
    Dim i As Double, j As Integer, k As Integer: j = 1
    Dim returnArray As Variant, obj As Variant
    
    For i = 2 To UBound(dataArray)
        On Error Resume Next
        uniqueCollection.Add Item:=dataArray(i, dimention), Key:=dataArray(i, dimention)
nxtUniqueItem:
    On Error GoTo 0
    Next i
    
    ReDim returnArray(0 To uniqueCollection.Count - 1)
    For i = 0 To UBound(returnArray)
        returnArray(i) = uniqueCollection(i + 1)
    Next i

    uniqueValuesFromRange = returnArray
End Function




Sub unHideAllSheets()
    Dim i As Integer
    For i = 1 To ActiveWorkbook.sheets.Count
        Worksheets(i).Visible = True
    Next i
End Sub
