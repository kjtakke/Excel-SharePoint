Sub ExampleAddButton()
    Call SetGrid("Sheet1")
    Call addButton("Sheet1", "1", "MyText", 5, 3, 2, 2)
End Sub

Sub SetGrid(ws As String)
    Worksheets(ws).Cells.ColumnWidth = 3.14
    Worksheets(ws).Cells.RowHeight = 15
End Sub

Sub addButton(ws As String, id As String, text As String, _
               x As Integer, y As Integer, w As Integer, h As Integer, _
               Optional macro As String = "", _
               Optional fontColor As MsoThemeColorIndex = msoThemeColorLight1, _
               Optional fillColor As MsoThemeColorIndex = msoThemeColorAccent1)
               
    Dim xv As Integer: xv = 20
    Dim yv As Integer: yv = 15
    
    With Worksheets(ws).Shapes.AddShape(msoShapeRectangle, (x * xv) + 1.4, y * yv, (w * xv) + 1.4, h * yv)
    '.Select
    .Name = "S" & id
    .TextFrame2.VerticalAnchor = msoAnchorMiddle
    .TextFrame2.TextRange.Characters.text = text
    
        With .TextFrame2.TextRange.Characters(1, 6).ParagraphFormat
            .FirstLineIndent = 0
            .Alignment = msoAlignCenter
        End With
        
        With .TextFrame2.TextRange.Characters(1, 6).Font
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.ObjectThemeColor = fontColor
            .Fill.ForeColor.TintAndShade = 0
            .Fill.ForeColor.Brightness = 0
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 11
            .Name = "+mn-lt"
        End With
        
        With .Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = fillColor
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    
        If macro <> "" Then .OnAction = macro
    End With
    
End Sub
