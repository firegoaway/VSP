Attribute VB_Name = "Module8"
Option Explicit

Private Const VERY_SMALL_NUMBER As Double = -1.79769313486231E+308

Sub FindLargestNumberAndReplacePlaceholder()
    Dim selectedTable As table
    Dim cell As cell
    Dim largestNumber As Double
    Dim cellValue As String
    Dim currentNumber As Double
    Dim i As Integer
    Dim columnNumber As Integer
    Dim doc As Document
    Dim rng As Range
    Dim nlines As Long
    
    ' Define the specific column number (1-based index)
    columnNumber = 4
    nlines = val(InputBox("¬веди 2, если в €чейке одно значение, 8, если в €чейке два значени€, 14, если значений три, 20, если значений четыре:", "nlines"))

    If Selection.Tables.Count = 0 Then
        MsgBox "Please select a table first!", vbExclamation
        Exit Sub
    End If

    Set selectedTable = Selection.Tables(1)
    largestNumber = VERY_SMALL_NUMBER  ' Initialize to the smallest possible number

    ' Iterate over rows in the specified column in the table
    On Error Resume Next
    For i = 1 To selectedTable.Rows.Count
        Set cell = selectedTable.cell(i, columnNumber)
        cellValue = cell.Range.text
        cellValue = Left(cellValue, Len(cellValue) - nlines)  ' Removing the end of cell marker
        If IsNumeric(cellValue) Then
            currentNumber = CDbl(cellValue)
            If currentNumber > largestNumber Then
                largestNumber = currentNumber
            End If
        End If
    Next

    On Error GoTo 0

    If largestNumber > VERY_SMALL_NUMBER Then
        Set doc = ActiveDocument
        Set rng = doc.Content
        rng.Find.ClearFormatting
        With rng.Find
            .text = "[[LARGEST_NUMBER_FROM_APPENDIX_1]]"
            .Replacement.text = CStr(largestNumber)
            .Execute Replace:=wdReplaceOne
        End With
        MsgBox "The largest number " & largestNumber & " in column " & columnNumber & " has replaced the placeholder in the document.", vbInformation
    Else
        MsgBox "No numeric values found in column " & columnNumber & ".", vbExclamation
    End If
End Sub

