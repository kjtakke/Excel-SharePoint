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


Sub add_task()
  Dim cnt As ADODB.Connection: Set cnt = New ADODB.Connection
  Dim rst As ADODB.Recordset: Set rst = New ADODB.Recordset
  Dim listID as String: listID = "7831A117-A687-4A6C-9945-7BD872C69BDC"
  Dim url as String: url = "http://SharePointSiteUrl"
  Dim tableName as string: tableName = "xxxxxxx"
  Dim mySql As String: mySql = "SELECT * FROM [" & tableName & "]"
  
  On Error GoTo en:

  With cnt
    .ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=" & url & ";LIST={" & listID & "};"
    .Open
  End With

  rst.Open mySql, cnt, adOpenDynamic, adLockOptimistic

  rst.AddNew
    rst.Fields("Field Name") = ""
    rst.Fields("Field Name") = ""
    rst.Fields("Field Name") = ""
    rst.Fields("Field Name") = ""
    rst.Fields("Field Name") = ""
    rst.Fields("Field Name") = ""
  rst.Update

  If CBool(rst.State And adStateOpen) = True Then rst.Close
  Set rst = Nothing
  If CBool(cnt.State And adStateOpen) = True Then cnt.Close
  Set cnt = Nothing

en:
End Sub

                
                
                
                
Sub add_Item(fileds As Variant, url As String, listID As String, ListName As String)
    'Const spURL = "https://xxxxxxx.sharepoint.com/teams/xxxxxxx"
    'Const spMyListID = "5AA6F12C-79CA-4454-B1B8-9FFB33CF996A"
    'Const spMyListName = "My List"
    'Sub AddSharePointListData()
    '    Dim fields(1 To 2, 1 To 2) As Variant
    '    fields(1, 1) = "Title"
    '    fields(2, 1) = "Description"
    '    fields(1, 2) = "Blad"
    '    fields(2, 2) = "Blah Blah"
    '    Call add_Item(fileds, spURL, spMyListID, spMyListName)
    'End Sub

    Dim cnt As ADODB.Connection: Set cnt = New ADODB.Connection
    Dim rst As ADODB.Recordset: Set rst = New ADODB.Recordset
    Dim mySql As String: mySql = "SELECT * FROM [" & ListName & "]"
    Dim i As Integer
    On Error GoTo en:
    
    With cnt
      .ConnectionString = _
      "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=" & url & ";LIST={" & listID & "};"
      .Open
    End With
    
    rst.Open mySql, cnt, adOpenDynamic, adLockOptimistic
    
    rst.AddNew
      For i = 1 To UBound(fields)
          rst.fields(fields(i, 1)) = fields(i, 2)
      Next i
    rst.Update
    
    If CBool(rst.State And adStateOpen) = True Then rst.Close
    Set rst = Nothing
    If CBool(cnt.State And adStateOpen) = True Then cnt.Close
    Set cnt = Nothing
    
en:
End Sub
