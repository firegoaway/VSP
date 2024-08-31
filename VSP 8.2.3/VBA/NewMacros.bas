Attribute VB_Name = "NewMacros"

Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос1"
'
' Макрос1 Макрос
'
'
    Selection.Font.Size = 12
    With Selection.ParagraphFormat
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitRightIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    With Selection.ParagraphFormat
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 1.15
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitRightIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
End Sub
Sub Макрос2()
Attribute Макрос2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос2"
'
' Макрос2 Макрос
'
'
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderLeft)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderRight)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderHorizontal)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderVertical)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub
Sub Укорачивание_текста_и_табуляции()
Attribute Укорачивание_текста_и_табуляции.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Укорачивание_текста_и_табуляции"
'
' Укорачивание_текста_и_табуляции Макрос
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Font.Size = 14
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 12
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 1.15
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
    End With
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Size = 16
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 14
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 1.15
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
    End With
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Size = 18
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 14
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 1.15
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
    End With
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub БыстраяЗамена()
Attribute БыстраяЗамена.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.БыстраяЗамена"
'
' БыстраяЗамена Макрос
'
'
    ActiveWindow.ActivePane.VerticalPercentScrolled = -54
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Windows( _
        "(РПР) 601204, Владимирская область, город Собинка, улица Димитрова, строение 20А.docx" _
        ).Activate
    Selection.Copy
    Selection.Copy
    Windows( _
        "(РПР) 601650, Владимирская область, район Александровский, город Александров, улица Ленина, дом 13.docx" _
        ).Activate
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=106.1, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=99, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=91.9, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=84.8, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=77.75, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=70.65, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=56.45, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=49.4, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=35.2, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=35.2, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=185.1, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=178, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=163.8, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=135.45, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=121.3, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=100.05, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=78.75, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=64.6, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=50.4, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=43.35, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=36.25, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=29.15, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=29.15, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).AllowAutoFit = False
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=190.6, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=174.75, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=160.55, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=132.2, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=118.05, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=103.85, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=89.7, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=82.6, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=68.45, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=61.35, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=51.05, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=51.05, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=84.8, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=77.75, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=63.55, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=56.45, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=49.4, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=42.3, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=35.2, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=28.1, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=21.05, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=13.95, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=13.95, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=128.4, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=114.2, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=92.95, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=71.7, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=50.4, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=36.25, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=22.05, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=15, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=11.55, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=11.55, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=125.15, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=110.95, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=89.7, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=68.45, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=47.15, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=25.9, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=11.75, RulerStyle:= _
        wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=11.75, RulerStyle:= _
        wdAdjustFirstColumn
    Windows( _
        "(РПР) 601204, Владимирская область, город Собинка, улица Димитрова, строение 20А.docx" _
        ).Activate
    Selection.Copy
    Selection.Copy
    Windows( _
        "(РПР) 601650, Владимирская область, район Александровский, город Александров, улица Ленина, дом 13.docx" _
        ).Activate
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.Copy
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "опасности здания: Ф3 "
        .Replacement.text = "опасности здания: Ф3.1 "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveDown Unit:=wdScreen, Count:=16
    Selection.MoveUp Unit:=wdScreen, Count:=1
    Selection.MoveDown Unit:=wdScreen, Count:=13
    Selection.MoveUp Unit:=wdScreen, Count:=2
    Selection.MoveDown Unit:=wdScreen, Count:=7
    Selection.MoveUp Unit:=wdScreen, Count:=2
    Selection.MoveDown Unit:=wdScreen, Count:=25
    ActiveWindow.ActivePane.VerticalPercentScrolled = -200
    Windows( _
        "(РПР) 601204, Владимирская область, город Собинка, улица Димитрова, строение 20А.docx" _
        ).Activate
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    Windows( _
        "(РПР) 601650, Владимирская область, район Александровский, город Александров, улица Ленина, дом 13.docx" _
        ).Activate
    ActiveWindow.ActivePane.VerticalPercentScrolled = -210
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
End Sub



Sub F3toF31()
    ' Define variables
    Dim findText As String
    Dim replaceText As String
    
    ' Set the text to find and replace
    findText = "опасности здания: Ф3 "
    replaceText = "опасности здания: Ф3.1 "
    
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



Sub ResizeImagesToPageWidth()
    Dim oShape As InlineShape
    Dim pageWidthCM As Single
    Dim shapeWidthInPoints As Single
    Dim pageWidthInPoints As Single
    
    ' Define the width of the page in centimeters
    pageWidthCM = 15
    
    ' Convert page width to points (1 cm = 28.35 points)
    pageWidthInPoints = pageWidthCM * 28.35
    
    ' Loop through each inline shape in the document
    For Each oShape In ActiveDocument.InlineShapes
        ' Get the width of the shape in points
        shapeWidthInPoints = oShape.Width
        
        ' Check if the image width is less than the page width
        If shapeWidthInPoints < pageWidthInPoints Then
            ' Maintain the aspect ratio and resize the image width to match the page width
            oShape.LockAspectRatio = msoTrue
            oShape.Width = pageWidthInPoints
        End If
    Next oShape
End Sub



Sub ResizeImagesWidthHeightCheck()
    Dim oShape As InlineShape
    Dim pageWidthCM As Single
    Dim pageHeightCM As Single
    Dim shapeWidthInPoints As Single
    Dim shapeHeightInPoints As Single
    Dim pageWidthInPoints As Single
    Dim pageHeightInPoints As Single
    Dim aspectRatioCM As Single
    
    ' Define the page width and height in centimeters
    pageWidthCM = val(InputBox("Enter Page Width (Default 15):", "Page Width"))
    pageHeightCM = val(InputBox("Enter Page Height (Default 23):", "Page Height"))
        
    ' Convert page width and height to points (1 cm = 28.35 points)
    pageWidthInPoints = pageWidthCM * 28.35
    pageHeightInPoints = pageHeightCM * 28.35
    
    ' Loop through each inline shape in the document
    For Each oShape In ActiveDocument.InlineShapes
        ' Get the width and height of the shape in points
        shapeWidthInPoints = oShape.Width
        shapeHeightInPoints = oShape.Height
        If shapeHeightInPoints - shapeWidthInPoints > 0 Then
            aspectRatioCM = shapeHeightInPoints / shapeWidthInPoints
        End If
        If shapeHeightInPoints - shapeWidthInPoints <= 0 Then
            aspectRatioCM = shapeWidthInPoints / shapeHeightInPoints
        End If
                
        ' Check if the image width is less than the page width and height is not greater than the page height
        If shapeWidthInPoints < pageWidthInPoints And shapeHeightInPoints <= pageHeightInPoints And aspectRatioCM < 2.25 Then
            ' Maintain the aspect ratio and resize the image width to match the page width
            oShape.LockAspectRatio = msoTrue
            oShape.Width = pageWidthInPoints
        End If
        If shapeWidthInPoints < pageWidthInPoints And shapeHeightInPoints <= pageHeightInPoints And aspectRatioCM >= 2.25 Then
            ' Maintain the aspect ratio and resize the image width to match the page width
            oShape.LockAspectRatio = msoTrue
            oShape.Width = pageWidthInPoints / (aspectRatioCM / (2.25 / 1.25))
        End If
    Next oShape
End Sub
Sub Level2Edit()
Attribute Level2Edit.VB_Description = "Приводит все заголовки стиля Level2 к единому виду, центрирует и делает 14 кеглем."
Attribute Level2Edit.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Level2Edit"
'
' Level2Edit Макрос
' Приводит все заголовки стиля Level2 к единому виду, центрирует и делает 14 кеглем.
'
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("LEVEL2")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 14
        .Bold = True
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 1.15
        .Alignment = wdAlignParagraphCenter
        .LineUnitBefore = 1.2
        .LineUnitAfter = 1.2
        .MirrorIndents = False
    End With
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub ЗаменаОпределенияНаРасчёт()
Attribute ЗаменаОпределенияНаРасчёт.VB_Description = "Заменяет заголовки на те, которые стыкуются с СП 505"
Attribute ЗаменаОпределенияНаРасчёт.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ЗаменаОпределенияНаРасчёт"
'
' ЗаменаОпределенияНаРасчёт Макрос
' Заменяет заголовки на те, которые стыкуются с СП 505
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Определение времени блокирования путей эвакуации"
        .Replacement.text = _
            "Расчёт времени блокирования путей эвакуации опасными факторами пожара"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Определение времени блокирования путей эвакуации"
        .Replacement.text = _
            "Расчёт времени блокирования путей эвакуации опасными факторами пожара"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Определение времени блокирования путей эвакуации"
        .Replacement.text = _
            "Расчёт времени блокирования путей эвакуации опасными факторами пожара"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = _
            "Составление расчётных схем и определение расчетного времени эвакуации людей"
        .Replacement.text = _
            "Составление расчётных схем эвакуации и расчет времени эвакуации людей"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = _
            "Составление расчётных схем и определение расчетного времени эвакуации людей"
        .Replacement.text = _
            "Составление расчётных схем эвакуации и расчет времени эвакуации людей"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = _
            "Составление расчётных схем и определение расчетного времени эвакуации людей"
        .Replacement.text = _
            "Составление расчётных схем эвакуации и расчет времени эвакуации людей"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Результаты моделирования движения людей"
        .Replacement.text = "Результаты моделирования эвакуации"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Результаты моделирования движения людей"
        .Replacement.text = "Результаты моделирования эвакуации"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Результаты моделирования движения людей"
        .Replacement.text = "Результаты моделирования эвакуации"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = _
            "Расчёт величины индивидуального пожарного риска для сценария"
        .Replacement.text = _
            "Расчёт величины индивидуального пожарного риска для человека в сценарии"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = _
            "Расчёт величины индивидуального пожарного риска для сценария"
        .Replacement.text = _
            "Расчёт величины индивидуального пожарного риска для человека в сценарии"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = _
            "Расчёт величины индивидуального пожарного риска для сценария"
        .Replacement.text = _
            "Расчёт величины индивидуального пожарного риска для человека в сценарии"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Расчёт величины индивидуального пожарного риска для здания"
        .Replacement.text = _
            "Расчёт величины индивидуального пожарного риска для человека на объекте"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub DeletePerechenIshodnih()
Attribute DeletePerechenIshodnih.VB_Description = "Удаление заголовка ""Перечень исходных данных"" из отчёта"
Attribute DeletePerechenIshodnih.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.DeletePerechenIshodnih"
'
' DeletePerechenIshodnih Макрос
' Удаление заголовка "Перечень исходных данных" из отчёта
'
    Selection.Find.Execute
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Перечень исходных данных^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
