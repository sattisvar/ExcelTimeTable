Private Sub UpdateButton_Click()
    Dim ModeRecalcul As Long
    ModeRecalcul = Application.Calculation
    Application.Calculation = xlCalculationManual ' bloque le calcul automatique.
    
    Dim dTime As Double
    dTime = MicroTimer
    Dim Plage As Range, c As Range, oneCell As Range
    Dim GetLastRowForThisWeek As Long, GetFirstRowForThisWeek As Long, LastRow As Long
    GetFirstRowForThisWeek = Worksheets("Annuel").Range("B1").End(xlDown).Row
    GetLastRowForThisWeek = Worksheets("Annuel").Range("B" & GetFirstRowForThisWeek).End(xlDown).Row
    
    ProgressBar.Show False
    Dim PercentVal As Double
    PercentVal = 0
    Dim PercentStep As Double
    LastRow = Worksheets("Annuel").Range("B100000").End(xlUp).Row
    PercentStep = 10 / (LastRow - GetFirstRowForThisWeek + 1 - WorksheetFunction.CountBlank(Worksheets("Annuel").Range("B" & GetFirstRowForThisWeek & ":B" & LastRow)))
    'Plage is the range of one week
    Set Plage = Worksheets("Annuel").Range("C" & GetFirstRowForThisWeek & ":L" & GetLastRowForThisWeek)
    
    Do While Plage.Cells(1, 1).Offset(0, -1).Value <> ""
    'We loop until we find the end of a week
    'which is checked by a Plage with no time stored in the column B
        For Each col In Plage.Columns
            Dim DayDate As Long
            DayDate = col.Cells(-1, 1).MergeArea.Cells(1, 1).Value
            For Each oneCell In col.Cells
                If (oneCell.Address = oneCell.MergeArea.Cells(1, 1).Address And oneCell <> "" And Worksheets("Annuel").Range("B" & oneCell.Row).Value <> "") Then
                    Dim cren As New Creneau
                    cren.Reset
                    cren.Lire oneCell, CDate(DayDate)
                    Dim uE As String
                    uE = cren.uE
                    With oneCell
                        .FormatConditions.Delete
                        For Each colorCell In [Listes!E3:E26]
                            If colorCell.Value = uE Then
                                .Interior.Color = colorCell.Interior.Color
                            End If
                        Next colorCell
                    End With
                End If
                PercentVal = PercentVal + PercentStep
            Next oneCell
            ProgressBar.Label1.Caption = Int(PercentVal) & "% completed"
            ProgressBar.Label1.Width = PercentVal * 3
            DoEvents
        Next col
        GetFirstRowForThisWeek = Worksheets("Annuel").Range("B" & CStr(GetLastRowForThisWeek + 1)).End(xlDown).Row
        GetLastRowForThisWeek = Worksheets("Annuel").Range("B" & GetFirstRowForThisWeek).End(xlDown).Row
        Set Plage = Worksheets("Annuel").Range("C" & GetFirstRowForThisWeek & ":L" & GetLastRowForThisWeek)
    Loop
    dTime = MicroTimer - dTime
    Unload ProgressBar
    Dim MsgT As String
    MsgT = "La mise à jour a réussi et a pris " & Int(dTime) & "s"
    MsgBox MsgT, vbOKOnly, "Infos"
    
    Application.Calculation = ModeRecalcul ' met en place le calcul automatique
End Sub
