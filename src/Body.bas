Sub Body()
Attribute Body.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Body"
'
' Body macro to format main body text into 
' IEEE two column report spec.
'
' Authors: David Suh
'
    Selection.Font.Name = "Times New Roman"
'   Creating Two Columns
    With ActiveDocument.PageSetup.TextColumns
        .SetCount NumColumns:= 2
        .EvenlySpaced = True
        .LineBetween = False
        .Width = InchesToPoints(3.5)
        .Spacing = InchesToPoints(0.25)
    End With
    
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(0.75)
        .BottomMargin = InchesToPoints(0.75)
        .LeftMargin = InchesToPoints(0.63)
        .RightMargin = InchesToPoints(0.63)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.5)
        .FooterDistance = InchesToPoints(0.5)
        .PageWidth = InchesToPoints(8.5)
        .PageHeight = InchesToPoints(11)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeDefault
    End With
End Sub
