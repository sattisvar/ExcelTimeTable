'Extraction of the Listing
Private Sub ExtractListing_Click()
    'oneCell will store the cells in one column of Plage
    'col will store one column of Plage
    Dim Plage As Range, oneCell As Range, col As Range, CountError As Integer
    Dim dTime As Double
    dTime = MicroTimer
    
    Dim GetLastRowForThisWeek As Long, GetFirstRowForThisWeek As Long, LastRow As Long
    GetFirstRowForThisWeek = Worksheets("Annuel").Range("B1").End(xlDown).Row
    GetLastRowForThisWeek = Worksheets("Annuel").Range("B" & GetFirstRowForThisWeek).End(xlDown).Row
    
    ProgressBar.Show False
    Dim PercentVal As Double
    PercentVal = 0
    Dim PercentStep As Double
    LastRow = Worksheets("Annuel").Range("B100000").End(xlUp).Row
    PercentStep = 10 / (LastRow - GetFirstRowForThisWeek + 1 - WorksheetFunction.CountBlank(Worksheets("Annuel").Range("B" & GetFirstRowForThisWeek & ":B" & LastRow)))
    CountError = 0
    'Plage is the range of one week
    Set Plage = Worksheets("Annuel").Range("C" & GetFirstRowForThisWeek & ":L" & GetLastRowForThisWeek)
    'iListingRow will store the row number of the next chunk to write in the listing
    Dim iListingRow As Long
    iListingRow = 2
    Worksheets("Erreurs").Range("A2:Z10000").Clear
    Do While Plage.Cells(1, 1).Offset(0, -1).Value <> ""
        'We loop until we find the end of a week
        'which is checked by a Plage with no time stored in the column B
        For Each col In Plage.Columns
            Dim DayDate As Long
            DayDate = col.Cells(-1, 1).MergeArea.Cells(1, 1).Value
            For Each oneCell In col.Cells
                If (oneCell.Address = oneCell.MergeArea.Cells(1, 1).Address And oneCell <> "" And Range("B" & oneCell.Row).Value <> "") Then
                    Dim cren As New Creneau
                    cren.Reset
                    cren.Lire oneCell, CDate(DayDate)
                    Dim WriteHere As Range
                    Dim eIdx As Integer
                    For eIdx = 0 To cren.HowManyEnseignants()
                        Set WriteHere = Worksheets("Listing").Range("A" & iListingRow)
                        Set HeadHere = Worksheets("Listing").Range("A1")
                        Do While HeadHere.Value <> ""
                            If (HeadHere.Value = Worksheets("Listes").Range("I3")) Then
                                WriteHere.Value = DayDate
                                WriteHere.NumberFormat = Worksheets("Listes").Range("J3")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I4")) Then
                                WriteHere.Value = DayDate
                                WriteHere.NumberFormat = Worksheets("Listes").Range("J4")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I5")) Then
                                'Contain the time of beginning
                                WriteHere.Value = cren.Beginning
                                WriteHere.NumberFormat = Worksheets("Listes").Range("J5")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I6")) Then
                                'Contain the time of the end of that class
                                WriteHere.Value = cren.Ending
                                WriteHere.NumberFormat = Worksheets("Listes").Range("J6")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I7")) Then
                                WriteHere.Value = cren.TimeDelta
                                WriteHere.NumberFormat = Worksheets("Listes").Range("J7")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I8")) Then
                                'Contain the name of the course
                                WriteHere.Value = oneCell
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I9")) Then
                                'Contain the UE. comparing with one of those in the liste
                                WriteHere.Value = cren.UE 'FindInRange(oneCell.Value, [Listes!E3:E26])
                                AddCommentInCellIfEmpty WriteHere, oneCell, ""
                                CountError = CountError + AddCommentInCellIfEmpty(WriteHere, oneCell, "UE non renseignée")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I10")) Then
                                'Contain the Subject. comparing with one of those in the liste
                                WriteHere.Value = cren.Discipline ' FindInRange(oneCell.Value, [Listes!C3:C28])
                                CountError = CountError + AddCommentInCellIfEmpty(WriteHere, oneCell, "Discipline non renseignée")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I11")) Then
                                'Contain the Teacher. comparing with one of those in the liste
                                WriteHere.Value = cren.GetEnseignant(eIdx) ' FindInRange(oneCell.Value, [Listes!A3:A85])
                                CountError = CountError + AddCommentInCellIfEmpty(WriteHere, oneCell, "Enseignant-e non renseigné-e")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I12")) Then
                                WriteHere.Value = "P" & Format(DayDate, "yymmdd") & Format(cren.Beginning, "hhmm")
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I14")) Then
                                WriteHere.Value = cren.WriteSalles()
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I15")) Then
                                WriteHere.Value = cren.Commentaire
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I16")) Then
                                If oneCell.Comment Is Nothing Then
                                    WriteHere.Value = ""
                                Else
                                    WriteHere.Value = oneCell.Comment.Text
                                End If
                            ElseIf (HeadHere.Value = Worksheets("Listes").Range("I17")) Then
                                Dim GroupeClass As String ' 1/1 pour classe entiere 1/2 ou 2/2 pour demi groupe 1/3 ou 2/3 ou 3/3 pour tiers de groupe, etc
                                If oneCell.MergeArea.Columns.Count = 1 Then
                                    If oneCell.MergeArea.Cells(1, 1).Column Mod 2 = 1 Then
                                        GroupeClass = "'1/2"
                                    Else
                                        GroupeClass = "'2/2"
                                    End If
                                Else
                                    GroupeClass = "'" + Trim(str(1 + eIdx)) + "/" + Trim(str(1 + cren.HowManyEnseignants))
                                End If
                                WriteHere.Value = GroupeClass
                            End If
                            Set WriteHere = WriteHere.Offset(0, 1)
                            Set HeadHere = HeadHere.Offset(0, 1)
                        Loop
                        'Going to next row for writing the next lesson
                        iListingRow = iListingRow + 1
                    Next eIdx
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
    MsgT = "L'extraction a réussi et a pris " & Int(dTime) & "s"
    If CountError <> 0 Then
        MsgT = MsgT & ", néanmoins " & CountError & " informations n'ont pas été trouvés; voir la feuille intitulée ""Errors"""
    Else
        MsgT = MsgT & ". Tout s'est bien passé"
    End If
    MsgBox MsgT, vbOKOnly, "Infos"
    If CountError <> 0 Then
        Application.Goto Reference:=Worksheets("Erreurs").Range("A1"), Scroll:=False
    End If
    Set Plage = Nothing
End Sub

Private Sub CreneauxLinesHeight()
    Dim GetLastRowForThisWeek As Long, GetFirstRowForThisWeek As Long
    GetLastRowForThisWeek = 0
    Do While GetLastRowForThisWeek < 10000
        GetFirstRowForThisWeek = Worksheets("Annuel").Range("B" & CStr(GetLastRowForThisWeek + 1)).End(xlDown).Row
        GetLastRowForThisWeek = Worksheets("Annuel").Range("B" & GetFirstRowForThisWeek).End(xlDown).Row
        Worksheets("Annuel").Rows(GetFirstRowForThisWeek & ":" & GetLastRowForThisWeek).RowHeight = 34.5
    Loop
End Sub
Private Function FindInRange(str As String, Plage As Range) As String
    Dim v As Variant
    v = Plage.Value2
    For j = LBound(v) To UBound(v)
        If (InStr(str, v(j, 1)) <> 0) Then
            FindInRange = v(j, 1)
            Exit Function
        End If
    Next j
End Function

Private Function AddCommentInCellIfEmpty(cellToTest As Range, rngToComment As Range, Comment As String) As Integer
    AddCommentInCellIfEmpty = 0
    'Dim BeginComment As String
    
    'If rngToComment.Comment Is Nothing And Comment <> "" And cellToTest = "" Then
    '    rngToComment.AddComment
    'End If
    'If Comment = "" And Not rngToComment.Comment Is Nothing Then
    '    rngToComment.Comment.Delete
    'End If
    If Comment <> "" And cellToTest = "" Then
        'BeginComment = rngToComment.Comment.Text
        'rngToComment.Comment.Text BeginComment & Chr(10) & Comment
        cellToTest.Interior.Color = RGB(255, 96, 96)
        
        If IsEmpty(Worksheets("Erreurs").Range("A1").Value) Then
            nextrow = 1
        Else
            nextrow = 1 + Worksheets("Erreurs").Range("A10000").End(xlUp).Row
        End If
        Bcol = cellToTest.Worksheet.Range("A1:Z1").Find(Worksheets("Listes").Range("I3").Value, LookIn:=xlValues).Column
        Ccol = cellToTest.Worksheet.Range("A1:Z1").Find(Worksheets("Listes").Range("I5").Value, LookIn:=xlValues).Column
        Dcol = cellToTest.Worksheet.Range("A1:Z1").Find(Worksheets("Listes").Range("I6").Value, LookIn:=xlValues).Column
        Worksheets("Erreurs").Range("A" & nextrow).Value = Comment
        Worksheets("Erreurs").Range("B" & nextrow).Value = _
                Format(cellToTest.Offset(0, Bcol - cellToTest.Column).Value, "dddd dd mmmm yyyy")
        Worksheets("Erreurs").Range("C" & nextrow).Value = _
                Format(cellToTest.Offset(0, Ccol - cellToTest.Column).Value, "h:mm")
        Worksheets("Erreurs").Range("D" & nextrow).Value = _
                Format(cellToTest.Offset(0, Dcol - cellToTest.Column).Value, "h:mm")
        Worksheets("Erreurs").Range("E" & nextrow).Value = rngToComment.Address
        Worksheets("Erreurs").Range("F" & nextrow).Value = Worksheets("Erreurs").Range("F1").Value
        Worksheets("Erreurs").Range("G" & nextrow).Value = cellToTest.Address
        Worksheets("Erreurs").Range("H" & nextrow).Value = Worksheets("Erreurs").Range("H1").Value
        
        AddCommentInCellIfEmpty = 1
    Else
        cellToTest.Interior.ColorIndex = 0
    End If
End Function


Private Sub InsertLesson_Click()
    On Error Resume Next
    CreneauFrm.Show
End Sub






