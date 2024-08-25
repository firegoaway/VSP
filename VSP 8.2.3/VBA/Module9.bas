Attribute VB_Name = "Module9"
Sub F5toF51()
    ' Define variables
    Dim findText As String
    Dim replaceText As String
    
    ' Set the text to find and replace
    findText = "опасности здания: Ф5 "
    replaceText = "опасности здания: Ф5.1 "
    
    ' Use the Word Find and Replace feature
    With ActiveDocument.Content.Find
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Execute Replace All
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub F5toF52()
    ' Define variables
    Dim findText As String
    Dim replaceText As String
    
    ' Set the text to find and replace
    findText = "опасности здания: Ф5 "
    replaceText = "опасности здания: Ф5.2 "
    
    ' Use the Word Find and Replace feature
    With ActiveDocument.Content.Find
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Execute Replace All
        .Execute Replace:=wdReplaceAll
    End With
End Sub


