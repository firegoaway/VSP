Attribute VB_Name = "Module6"
Sub StandardizeTables()
    Dim tbl As table
    Dim cel As cell

    For Each tbl In ActiveDocument.Tables
        For Each cel In tbl.Range.Cells
            With cel.Range.Font
                .Name = "Times New Roman"
                .Size = 12
            End With
            With cel.Range.ParagraphFormat
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
            End With
        Next cel

        With tbl.Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
        End With
    Next tbl
End Sub

