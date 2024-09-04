Attribute VB_Name = "Module2"
Sub ReplacePhrase()
    Dim targetDoc As Document
    Set targetDoc = ActiveDocument
    With targetDoc.Content.Find
        .Text = "приказ МЧС РФ от 30.06.2009 № 382"  ' The phrase you want to search for"
        .Replacement.Text = "[4]"  ' The text you want to replace it with
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll  ' Execute the replacement operation
    End With
End Sub
