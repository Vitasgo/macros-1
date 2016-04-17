Sub Body()
Attribute Body.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Body"
'
' Body macro to format main body text into 
' IEEE two column report spec for keywords.
'
' Authors: David Suh
'
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 10
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Or ActiveWindow.ActivePane.View.Type _
         = wdMasterView Then
        ActiveWindow.ActivePane.View.Type = wdPageView
    End If
    With ActiveDocument.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = False
        .Width = InchesToPoints(3.5)
        .Spacing = InchesToPoints(0.25)
    End With
End Sub
