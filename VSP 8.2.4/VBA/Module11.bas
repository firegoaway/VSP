Attribute VB_Name = "Module11"
Sub DeleteRowsWithSpecificText()
    Dim tbl As table
    Dim cel As cell
    Dim doc As Document
    Dim found As Boolean

    Set doc = ActiveDocument
    found = False
    
    For Each tbl In doc.Tables
        For Each cel In tbl.Range.Cells
            If InStr(cel.Range.text, "не используется") > 0 Then
                cel.Select
                found = True
                Selection.Rows.Delete
                Exit For
            End If
        Next cel
        If found Then Exit For
    Next tbl
    
    If Not found Then
        MsgBox "Text not found in any cell."
    End If
End Sub

