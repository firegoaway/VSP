Attribute VB_Name = "Module5"
Option Explicit

' Windows API functions to interact with the clipboard
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As String) As LongPtr

' Clipboard constants
Private Const GMEM_MOVEABLE As Long = &H2
Private Const CF_TEXT As Long = 1

Sub FindLargestNumberAndCopyToClipboard()
    Dim selectedTable As table
    Dim cell As cell
    Dim largestNumber As Double
    Dim cellValue As String
    Dim currentNumber As Double
    Dim i As Integer
    Dim columnNumber As Integer
    Dim nlines As Long

    ' Change this to target the specific column number (1-based index)
    columnNumber = 4
    nlines = val(InputBox("¬веди 2, если в €чейке одно значение, 8, если в €чейке два значени€, 14, если значений три, 20, если значений четыре:", "nlines"))

    If Selection.Tables.Count = 0 Then
        MsgBox "Please select a table first!", vbExclamation
        Exit Sub
    End If

    Set selectedTable = Selection.Tables(1)
    largestNumber = -1.79769313486231E+308  ' Initialize

    On Error Resume Next  ' Avoid errors with non-numeric values
    ' Iterate over each cell in the specified column
    For i = 1 To selectedTable.Rows.Count
        Set cell = selectedTable.cell(i, columnNumber)
        cellValue = cell.Range.text
        cellValue = Left(cellValue, Len(cellValue) - nlines)  ' Strip end-of-cell character
        If IsNumeric(cellValue) Then
            currentNumber = CDbl(cellValue)
            If currentNumber > largestNumber Then
                largestNumber = currentNumber
            End If
        End If
    Next

    On Error GoTo 0

    If largestNumber > -1.79769313486231E+308 Then
        CopyTextToClipboard CStr(largestNumber)
        MsgBox "The largest number " & largestNumber & " in column " & columnNumber & " has been copied to the clipboard.", vbInformation
    Else
        MsgBox "No numeric values found in column " & columnNumber & ".", vbExclamation
    End If
End Sub

Private Sub CopyTextToClipboard(text As String)
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr

    ' Allocate moveable memory
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, Len(text) + 1)
    If hGlobalMemory <> 0 Then
        lpGlobalMemory = GlobalLock(hGlobalMemory)
        If lpGlobalMemory <> 0 Then
            lstrcpy lpGlobalMemory, text
            GlobalUnlock hGlobalMemory
            
            ' Open Clipboard to copy data
            If OpenClipboard(0&) Then
                EmptyClipboard
                SetClipboardData CF_TEXT, hGlobalMemory
                CloseClipboard
            End If
        End If
    End If
End Sub

