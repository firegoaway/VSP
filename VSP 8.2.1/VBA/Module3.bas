Attribute VB_Name = "Module3"
Sub InsertTextFromSpecificPages()
    Dim sourceDocument As Document
    Dim sourcePath As String
    Dim startPage As Long
    Dim endPage As Long
    Dim currentPage As Long
    Dim contentToCopy As Range
    Dim startRange As Range
    Dim endRange As Range

    ' Updating the status of the visual elements on the screen
    Application.ScreenUpdating = False

    ' Set the path to the source document
    sourcePath = "E:\Downloads\AnalysisF31.docx"
        
    ' Prompt the user for the starting and ending page numbers
    startPage = Val(InputBox("Enter the starting page number:", "Start Page"))
    endPage = Val(InputBox("Enter the ending page number:", "End Page"))

    If startPage <= 0 Or endPage < startPage Then
        MsgBox "Invalid page range. Please enter valid starting and ending page numbers.", vbExclamation
        Application.ScreenUpdating = True ' Re-enable the screen updates
        Exit Sub
    End If

    ' Attempt to open the source document
    Set sourceDocument = Documents.Open(FileName:=sourcePath, ReadOnly:=True, Visible:=False)
    If sourceDocument Is Nothing Then
        MsgBox "The document could not be opened.", vbExclamation
        Application.ScreenUpdating = True ' Re-enable the screen updates
        Exit Sub
    End If

    ' Define the initial range to copy
    Set startRange = sourceDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=startPage)
    startRange.Collapse Direction:=wdCollapseStart

    ' Loop until we find the end of the desired page range
    Set endRange = startRange.Duplicate
    For currentPage = startPage To endPage
        Set endRange = sourceDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=currentPage + 1)
        endRange.Collapse Direction:=wdCollapseStart
    Next currentPage
    
    ' Go back one character from the end of the last page to correctly define the range
    endRange.MoveEnd Unit:=wdCharacter, Count:=-1

    ' Combine into one range from startRange's start to endRange's end
    Set contentToCopy = sourceDocument.Range(Start:=startRange.Start, End:=endRange.End)

    ' Copy the content we want
    contentToCopy.Copy

    ' Paste into the selection of the active document
    Selection.PasteAndFormat (wdFormatOriginalFormatting)

    ' Close the source document
    sourceDocument.Close SaveChanges:=wdDoNotSaveChanges

    ' Re-enable the screen updates
    Application.ScreenUpdating = True
End Sub
