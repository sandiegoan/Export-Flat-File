Option Compare Database

Sub ExportFile()
    Dim strSQL As String
    Dim rs As DAO.Recordset
    Dim text_file As String
    
    strSQL = "SELECT * FROM [Table Name]"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    text_file = "C:\Output File Path.txt"
    
    'Open output file
    
    Open text_file For Output As #1
    
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        While (Not rs.EOF)
            Print #1, rs.Fields("Field 1 Name")
            Print #1, rs.Fields("Field 1 Name")
            Print #1, rs.Fields("Field 1 Name")
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
    
    'Close output file
    Close #1
    
End Sub
