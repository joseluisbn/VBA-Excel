Private Sub

Worksheet_Change (ByVal Target As Range)
If Intersect (Target, Range("A1:C1000")) Is Nothing Or Target.Cells.Count > 1 Then Exit Sub
Application.EnableEvents = False
Target.Value = UCase (Target.Value)
Application.EnableEvents = True

End Sub
