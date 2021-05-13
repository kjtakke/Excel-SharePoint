  
'REQUIRED REFERENCES
    'VISUAL BASIC FOR APPLICATION
    'MICROSOFT EXCEL 16.0 OBJECT LIBRARY
    'OLE AUTOMATION
    'MICROSOFT OFFICE 16.0 OBJECT LIBRARY
    'MICROSOFT ACTIVEX DAATA OBJECTS 6.1 LIBRARY
    'MICROSOFT ACTIVEX DATA OBJECTS RECORDSET 6.0 LIBRARY
    'MICROSOFT SCRIPTING RUNTIME
    'MICROSOFT OUTLOOK 16.0 OBJECT LIBRARY
    'MICROSOFT FORMS 2.0 OBJECT LIBRARY

Const delim = ";"
Const lnBrk = "\n"
Const ext = ".txt"

Private Sub ImportFile()
    Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
    Dim fd As Office.FileDialog, strFile As String: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim outputArray As Variant, lines As Variant, cols As Variant, i As Double, j As Integer, k As Integer
    On Error GoTo exitFileSub:
    
    'Get file path
    With fd
        .Filters.Clear
        .Filters.Add "Excel Files", "*" & ext, 1
        .Title = "Choose an Excel file"
        .AllowMultiSelect = False
        .InitialFileName = "C:\VBA Folder"
        If .Show = True Then
            strFile = .SelectedItems(1)
        End If
    End With
    
    'Get file text
    Open strFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline
    Loop
    Close #1

    'Extract text into an array
    lines = Split(text, lnBrk)
    ReDim outputArray(1 To UBound(lines) + 1, 1 To 6)
    For i = 0 To UBound(lines)
        cols = Split(lines(i), delim)
        For j = 0 To UBound(cols)
            outputArray(i + 1, j + 1) = cols(j)
        Next j
    Next i
    
    'Place on sheet
    Application.ScreenUpdating = False
        Worksheets("DataModelFileData").range("A2:F100000").Value = ""
        Worksheets("DataModelFileData").range("A2:F" & UBound(outputArray)).Value = outputArray
    Application.ScreenUpdating = True
    
    Exit Sub
exitFileSub:
End Sub





Sub importErpFiles()
    skip = False
    Dim GetFolder As String, fldr As FileDialog, sItem As String, i As Double, j As Integer, h As Integer, k As Double:: k = 1
    Dim fileName As String, TextFile As Integer:  TextFile = 1
    Dim export As String, output As Variant: export = ""
    Dim arrayCollection As Collection: Set arrayCollection = New Collection
    Dim lines As Variant, cols As Variant
    
    'Folder Picker
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then
        skip = True
        GoTo NextCode
        End If
        sItem = .SelectedItems(1)
    End With
    
    'Loop through all files in a folder
    Dim fileNames As Variant
    fileNames = Dir(sItem & "\")
    
    While fileNames <> ""
        'Get file text
        If Right(fileNames, 4) = ext Then
            Open sItem & "\" & fileNames For Input As #1
            Do Until EOF(1)
                Line Input #1, textline
                text = text & textline
            Loop
            Close #1
            
            'Extract text into an array
            lines = Split(text, lnBrk)
            ReDim outputArray(1 To UBound(lines) + 1, 1 To 6)
            For i = 0 To UBound(lines)
                cols = Split(lines(i), delim)
                For j = 0 To UBound(cols)
                    outputArray(i + 1, j + 1) = cols(j)
                Next j
            Next i
            k = k + i

            'Add array to a collection
            arrayCollection.Add outputArray
            fileNames = Dir
        End If
    Wend
    
    'Fill array
    ReDim outputArray(1 To k, 1 To 6)
    k = 1
    For j = 1 To arrayCollection.Count
        For i = 1 To UBound(arrayCollection(j))
            For h = 1 To 6
                outputArray(k, h) = arrayCollection(j)(i, h)
            Next h
            k = k + 1
        Next i
    Next j
    
    'Place on sheet
    Application.ScreenUpdating = False
        Worksheets("DataModelFileData").range("A2:F100000").Value = ""
        Worksheets("DataModelFileData").range("A2:F" & UBound(outputArray)).Value = outputArray
    Application.ScreenUpdating = True
    
NextCode:
End Sub




Private Sub ExportFile_FromBaseOneArray()
    Dim GetFolder As String, fldr As FileDialog, sItem As String, i As Double, j As Integer
    Dim fileName As String, TextFile As Integer:  TextFile = 1
    Dim export As String, output As Variant: export = ""
    'Folder Picker
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
    'File name inputbox
    fileName = InputBox("FileName")
    On Error GoTo NextCode:
    GetFolder = sItem & "\" & fileName & ext
    Set fldr = Nothing
    
    'Create File
    output = Worksheets("LoadedList").range("A1").CurrentRegion.Value
    
    For i = 2 To UBound(output)
        For j = 1 To 6
            If j = 6 Then
                export = export & output(i, j) & lnBrk
            Else
                export = export & output(i, j) & delim
            End If
        Next j
    Next i
    
    'Write File
    Open GetFolder For Output As TextFile
        Print #TextFile, export
    Close TextFile
NextCode:
End Sub
