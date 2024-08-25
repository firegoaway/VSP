Attribute VB_Name = "Module4"
Sub DeleteSpecificRows()
    Dim tbl As Table
    Dim rw As Integer
    Dim cl As Integer
    Dim keyword1 As String
    Dim keyword2 As String
    Dim keyword3 As String
  
    ' Define the keywords
    keyword1 = "Ўирина"
    keyword2 = "”ровень нижнего по€са"
    keyword3 = "»меетс€ аварийна€ вентил€ци€"
  
    ' Loop through each table in the document
    For Each tbl In ActiveDocument.Tables
        ' Start from the last row to avoid index out-of-range errors after deleting rows
        For rw = tbl.Rows.Count To 2 Step -1
            ' Check each cell in the row
            For cl = 1 To tbl.Rows(rw).Cells.Count
                If InStr(tbl.Rows(rw).Cells(cl).Range.Text, keyword1) > 0 Or _
                   InStr(tbl.Rows(rw).Cells(cl).Range.Text, keyword2) > 0 Or _
                   InStr(tbl.Rows(rw).Cells(cl).Range.Text, keyword3) > 0 Then
                    ' If the keyword is found, delete this row and the one above it
                    tbl.Rows(rw).Delete
                    tbl.Rows(rw - 1).Delete
                    ' Once rows are deleted, exit the cell loop to avoid further processing on deleted rows
                    Exit For
                End If
            Next cl
        Next rw
    Next tbl
End Sub

