'Example
'
'Dim myObject As New Events
'
'Sub EventsOn()
'   Set myObject.SheetEvent = Application
'End Sub
'Sub EventsOff()
'   Set myObject.SheetEvent = Nothing
'End Sub


Public WithEvents SheetEvent As Application

Private Sub SheetEvent_SheetActivate(ByVal Sh As Object)
    Dim sn As String: sn = Sh.Name
    MsgBox ("You selected" & sn)
End Sub

Private Sub SheetEvent_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    MsgBox "You Made a change"
End Sub

Private Sub SheetEvent_WindowResize(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    MsgBox "you resized a window"
End Sub
