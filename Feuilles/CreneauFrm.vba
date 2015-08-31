

Public cren As New Creneau

Private Sub AddRoomToSelCmdBttn_Click()
    Dim iToRemove() As Integer
    Dim iToRemoveLen As Integer
    cren.VideSalles
    For intRow = 0 To RoomLstBox.ListCount - 1
        If RoomLstBox.Selected(intRow) Then
            SelectedRoomLstBox.AddItem RoomLstBox.List(intRow)
            ReDim Preserve iToRemove(0 To iToRemoveLen) As Integer
            iToRemove(iToRemoveLen) = intRow
            iToRemoveLen = iToRemoveLen + 1
        End If
    Next intRow
    For idxToRem = UBound(iToRemove) To LBound(iToRemove) Step -1
        RoomLstBox.RemoveItem iToRemove(idxToRem)
    Next idxToRem
    For intRow = 0 To SelectedRoomLstBox.ListCount - 1
        cren.AjouteSalle SelectedRoomLstBox.List(intRow)
    Next intRow
    FinalResultTxtBx.Text = cren.WriteStr()
End Sub


Private Sub AddTeacherToSelCmdBttn_Click()
    Dim iToRemove() As Integer
    Dim iToRemoveLen As Integer
    cren.VideEnseignants
    For intRow = 0 To TeachersLstBox.ListCount - 1
        If TeachersLstBox.Selected(intRow) Then
            SelectedTeachersLstBox.AddItem TeachersLstBox.List(intRow)
            ReDim Preserve iToRemove(0 To iToRemoveLen) As Integer
            iToRemove(iToRemoveLen) = intRow
            iToRemoveLen = iToRemoveLen + 1
        End If
    Next intRow
    For idxToRem = UBound(iToRemove) To LBound(iToRemove) Step -1
        TeachersLstBox.RemoveItem iToRemove(idxToRem)
    Next idxToRem
    For intRow = 0 To SelectedTeachersLstBox.ListCount - 1
        cren.AjouteEnseignant SelectedTeachersLstBox.List(intRow)
    Next intRow
    FinalResultTxtBx.Text = cren.WriteStr()
End Sub

Private Sub CancelBttn_Click()
    Unload Me
End Sub

Private Sub CommentsTxtBx_Change()
    cren.Commentaire = CommentsTxtBx.Value
    FinalResultTxtBx.Text = cren.WriteStr()
End Sub

Private Sub DisciplineCmbBx_Change()
    cren.Discipline = DisciplineCmbBx.Value
    FinalResultTxtBx.Text = cren.WriteStr()
End Sub

Private Sub EnseignantCmbBx_Change()
    cren.Enseignant = EnseignantCmbBx.Value
    FinalResultTxtBx.Text = cren.WriteStr()
End Sub

Private Sub ModifyBttn_Click()
    Selection.Cells(1, 1) = FinalResultTxtBx.Text
    With Selection
        .borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        .borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        .borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        .borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
    End With
    Unload Me
End Sub

Private Sub RemoveRoomFromSelCmdBttn_Click()
    Dim iToRemove() As Integer
    Dim iToRemoveLen As Integer
    cren.VideSalles
    For intRow = 0 To SelectedRoomLstBox.ListCount - 1
        If SelectedRoomLstBox.Selected(intRow) Then
            RoomLstBox.AddItem SelectedRoomLstBox.List(intRow)
            ReDim Preserve iToRemove(0 To iToRemoveLen) As Integer
            iToRemove(iToRemoveLen) = intRow
            iToRemoveLen = iToRemoveLen + 1
        End If
    Next intRow
    For idxToRem = UBound(iToRemove) To LBound(iToRemove) Step -1
        SelectedRoomLstBox.RemoveItem iToRemove(idxToRem)
    Next idxToRem
    For intRow = 0 To SelectedRoomLstBox.ListCount - 1
        cren.AjouteSalle SelectedRoomLstBox.List(intRow)
    Next intRow
    FinalResultTxtBx.Text = cren.WriteStr()
End Sub

Private Sub RemoveTeacherFromSelCmdBttn_Click()
    Dim iToRemove() As Integer
    Dim iToRemoveLen As Integer
    cren.VideEnseignants
    For intRow = 0 To SelectedTeachersLstBox.ListCount - 1
        If SelectedTeachersLstBox.Selected(intRow) Then
            TeachersLstBox.AddItem SelectedTeachersLstBox.List(intRow)
            ReDim Preserve iToRemove(0 To iToRemoveLen) As Integer
            iToRemove(iToRemoveLen) = intRow
            iToRemoveLen = iToRemoveLen + 1
        End If
    Next intRow
    For idxToRem = UBound(iToRemove) To LBound(iToRemove) Step -1
        SelectedTeachersLstBox.RemoveItem iToRemove(idxToRem)
    Next idxToRem
    For intRow = 0 To SelectedTeachersLstBox.ListCount - 1
        cren.AjouteEnseignant SelectedTeachersLstBox.List(intRow)
    Next intRow
    FinalResultTxtBx.Text = cren.WriteStr()
End Sub

Private Sub SelectedRoomLstBx_Click()

End Sub

Private Sub UECmbBx_Change()
    cren.UE = UECmbBx.Value
    FinalResultTxtBx.Text = cren.WriteStr()
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler:
    Set WsAnnuel = Worksheets("Annuel")
    Dim dayd As Date
    Set ThisCell = Selection.Cells(1, 1)
    With ThisCell
        DateRow = .Offset(0, 1 - .Column).End(xlUp).Row
        thisRow = .Row
        dayd = CDate(.Offset(DateRow - thisRow, 0))
        If dayd = 0 Then
            dayd = CDate(.Offset(DateRow - thisRow, -1))
        End If
        If (.Offset(0, 2 - .Column).Value = "") Then
            GoTo ErrHandler:
        End If
        Label3.Caption = "Créneau du " & Format(dayd, "ddd d mmm yyyy") _
                    & " de " & Format(.Offset(0, 2 - .Column), "hh:mm") _
                    & " à " & Format(.Offset(0, 2 - .Column).Offset(.MergeArea.Rows.Count - 1, 0), "hh:mm")
    End With
    '.Offset(-2, 0).Value
    FinalResultTxtBx.Text = Selection.Cells(1, 1)
    cren.Lire Selection.Cells(1, 1), dayd
    
    Set Ws = Sheets("Listes") 'Correspond au nom de votre onglet dans le fichier Excel
    For j = 3 To Ws.Range("A" & Rows.Count).End(xlUp).Row
        ThisEnseign = Ws.Range("A" & j)
        If cren.Enseignants(CStr(ThisEnseign)) Then
            Me.SelectedTeachersLstBox.AddItem ThisEnseign
        Else
            Me.TeachersLstBox.AddItem ThisEnseign
        End If
    Next j
    With Me.DisciplineCmbBx
        For j = 3 To Ws.Range("C" & Rows.Count).End(xlUp).Row
            ThisDiscipl = Ws.Range("C" & j)
            .AddItem ThisDiscipl
            If ThisDiscipl = cren.Discipline Then
                .Value = cren.Discipline
            End If
        Next j
    End With
    With Me.UECmbBx
        For j = 3 To Ws.Range("E" & Rows.Count).End(xlUp).Row
            ThisUE = Ws.Range("E" & j)
            .AddItem ThisUE
            If ThisUE = cren.UE Then
                .Value = cren.UE
            End If
        Next j
    End With
    For j = 3 To Ws.Range("D" & Rows.Count).End(xlUp).Row
        ThisRoom = Ws.Range("D" & j)
        If cren.SalleReservee(CStr(ThisRoom)) Then
            Me.SelectedRoomLstBox.AddItem ThisRoom
        Else
            Me.RoomLstBox.AddItem ThisRoom
        End If
    Next j
    CommentsTxtBx.Text = cren.Commentaire
    Exit Sub
ErrHandler:
    On Error Resume Next
    MsgBox "Soyez sur de sélectionner une plage de l'emploi du temps valide avant de lancer cette macro."
    Unload Me
End Sub
