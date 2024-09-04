Attribute VB_Name = "Module1"

Sub ResizeImages()
    Dim img As InlineShape
    Dim targetWidth As Single
    Dim targetHeight As Single
    Dim scaleFactor As Single
    
    ' User inputs for width and height
    targetWidth = InputBox("Enter the target width for the images (in points):", "Image Resizer", 100)
    If targetWidth <= 0 Then
        MsgBox "Invalid width input.", vbExclamation, "Error"
        Exit Sub
    End If
    
    targetHeight = InputBox("Enter the target height for the images (in points):", "Image Resizer", 100)
    If targetHeight <= 0 Then
        MsgBox "Invalid height input.", vbExclamation, "Error"
        Exit Sub
    End If
    
    For Each img In ActiveDocument.InlineShapes
        If img.Type = wdInlineShapePicture Then
            ' Calculate the scale factor based on the desired width and current width to maintain aspect ratio
            scaleFactor = targetWidth / img.Width
            
            ' Prevent upscaling if the target size is bigger than the original image size
            If scaleFactor < 1 Then
                img.ScaleWidth = scaleFactor * 100
                img.ScaleHeight = scaleFactor * 100
                
                ' Check and adjust height if it exceeds the user input after scaling proportionally
                If img.Height > targetHeight Then
                    scaleFactor = targetHeight / img.Height
                    img.ScaleHeight = scaleFactor * 100
                    img.ScaleWidth = scaleFactor * 100
                End If
            End If
        End If
    Next img
    
    MsgBox "Images have been resized.", vbInformation, "Done"
End Sub
