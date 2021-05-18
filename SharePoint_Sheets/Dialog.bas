'Vairables, objects and constants
'############################################################
Const xl = 68.25
Const yt = 30.75




'Testing
'############################################################
Sub TestAll()
    'Sheet1
    'id     Title
    '1      Task 1
    '2      Task 2
    '3      Task 3
    '4      Task 4
    '5      Task 5
    '6      Task 6
    '7      Task 7
    '8      Task 8
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        On Error Resume Next
            Sheets("Gallery").Delete
        On Error GoTo 0
    Application.DisplayAlerts = True
    TestDialog
    Dim i As Integer, j As Integer: j = 2
    Dim taskArray As Variant: taskArray = Worksheets("Sheet1").Range("A1").CurrentRegion.Value
    
    For i = 10 To UBound(taskArray) * 20
        CreateLabel ws:="Gallery", objName:="Lbl" & taskArray(j, 1) & j, text:=taskArray(j, 2) & "", x:=30, y:=10 + i, w:=300, h:=15
        i = i + 19
        
        j = j + 1
        If j > UBound(taskArray) Then Exit For
    Next i
    
    j = 2
    For i = 10 To UBound(taskArray) * 20
        CreateButton macro:="TestMessage", ws:="Gallery", objName:="Btn" & taskArray(j, 1) & j, text:=">", x:=5, y:=10 + i, w:=20, h:=15
        i = i + 19
        j = j + 1
        If j > UBound(taskArray) Then Exit For
    Next i
    Worksheets("Sheet1").Select
    Sheets("Gallery").Visible = False
    TestFormShow
    
    Application.ScreenUpdating = True
End Sub

Sub TestMessage()
    MsgBox ("Button Worked")
End Sub
Sub testingSubs()
    CreateLabel ws:="Gallery", objName:="Lbl1", text:="Blah", x:=50, y:=10, w:=100, h:=15
End Sub

Sub TestFormShow()
    ShowDialogForm form:="Gallery"
End Sub

Sub TestDialog()
    CreateDiologSheet Name:="Gallery", caption:="My Form"
End Sub

Sub testingButtons()
    CreateButton ws:="Gallery", objName:="Lbl1", text:="Blah", x:=50, y:=10, w:=100, h:=15
End Sub

'Main
'############################################################
Sub CreateButton(macro As String, ws As String, objName As String, text As String, x As Integer, y As Integer, Optional w As Integer = 100, Optional h As Integer = 15)
    With Sheets(ws)
        .Buttons.Add(xl + x, yt + y, w, h).Name = objName
        .Buttons(objName).Characters.text = text
        .Buttons(objName).OnAction = macro
    End With
End Sub

Sub CreateLabel(ws As String, objName As String, text As String, x As Integer, y As Integer, Optional w As Integer = 100, Optional h As Integer = 15)
    With Sheets(ws)
        .Labels.Add(xl + x, yt + y, w, h).Name = objName
        .Labels(objName).Characters.text = text
    End With
End Sub


Sub CreateDiologSheet(Name As String, caption As String)
    Application.ScreenUpdating = False
        DialogSheets.Add
        ActiveSheet.Name = Name
        Sheets(Name).Shapes("Dialog 1").Select
        Selection.Characters.text = "Dialog Caption"
        Selection.ShapeRange.IncrementLeft 0
        Selection.ShapeRange.IncrementTop 0
        Selection.ShapeRange.ScaleWidth 2.46, msoFalse, msoScaleFromTopLeft
        Selection.ShapeRange.ScaleHeight 3.61, msoFalse, msoScaleFromTopLeft
        Selection.Characters.text = caption
        Sheets(Name).Shapes("Button 2").Delete
        Sheets(Name).Shapes("Button 3").Delete
    Application.ScreenUpdating = True
End Sub


Sub ShowDialogForm(form As String)
    DialogSheets(form).Show
End Sub


Function GetButtonName(ws As String) As String
    Dim Obj As Object
    Dim ObjType As String
    On Error Resume Next
    Set Obj = Sheets(ws).Buttons(Application.Caller)
    If Err = 0 Then
        ObjType = "Button"
    Else
        Set Obj = Sheets(ws).Shapes(Application.Caller)
        ObjType = "Shape"
    End If
    On Error GoTo 0
    Select Case ObjType
        Case "Button"
            GetButtonName = Obj.Name
        Case "Shape"
            GetButtonName = Obj.Name
    End Select
End Function
