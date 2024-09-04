Attribute VB_Name = "Module7"
Sub ReplaceKoshmarov()
    Dim findText As String
    Dim replaceText As String
    Dim myRange As Range
    
    ' Set the findText and replaceText variables with the specified phrases
    findText = "�������� �. �. ��������������� ������� �������� ������ � ���������: ������� �������. � �.: �������� ��� ��� ������, 2000. � 118 �"
    replaceText = "�������� �. �., ����� �. �., ��������������� ������� �������� ������ � ���������: ������� �������. � �.: �������� ��� ��� ������, 2012. � 121 �"

    ' Initialize the Range object to represent the entire document
    Set myRange = ActiveDocument.Content
    
    ' Use the Find and Replace methods
    With myRange.Find
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    ' Execute the Find and Replace operation
    myRange.Find.Execute Replace:=wdReplaceAll
End Sub

