Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Count = 1 And Target.Row <> 1 Then
        If Target.Column = 6 Then
            If Target.Offset(0, -1).Value <> "" Then
                Application.Goto Reference:=Worksheets("Annuel").Range(Target.Offset(0, -1).Value), _
                        Scroll:=False
            End If
        ElseIf Target.Column = 8 Then
            If Target.Offset(0, -1).Value <> "" Then
                Application.Goto Reference:=Worksheets("Listing").Range(Target.Offset(0, -1).Value), _
                        Scroll:=False
            End If
        End If
    End If
End Sub




