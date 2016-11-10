Private enseignant_() As String
Private ue_ As String
Private discipline_ As String
Private salles_() As String
Private commentaires_ As String
Private DateHourOfBegin As Date
Private DateHourOfEnd As Date
Private Attached_Worksheet As Worksheet
Private TimeLength As Double
Private Valid_ As Boolean
Public Function HowManyEnseignants() As Integer
    If (Not Not enseignant_) Then
        HowManyEnseignants = UBound(enseignant_) - LBound(enseignant_)
    Else
        HowManyEnseignants = 0
    End If
End Function

Public Function Validity() As Boolean
    Validity = Valid_
End Function
Public Function Beginning() As Double
    Beginning = DateHourOfBegin - Int(DateHourOfBegin)
End Function

Public Function Ending() As Double
    Ending = DateHourOfEnd - Int(DateHourOfEnd)
End Function

Public Function TimeDelta() As Double
    TimeDelta = TimeLength
End Function

Public Property Let Commentaire(strCom As String)
    commentaires_ = strCom
End Property

Public Property Get Commentaire() As String
    Commentaire = commentaires_
End Property

Public Property Let Discipline(strDis As String)
    discipline_ = strDis
End Property

Public Property Get Discipline() As String
    Discipline = discipline_
End Property

Public Property Let UE(strUe As String)
    ue_ = strUe
End Property

Public Property Get UE() As String
    UE = ue_
End Property

Public Sub AjouteSalle(oneSalle As String)
    If (Not Not salles_) Then
        For i = LBound(salles_) To UBound(salles_)
            If salles_(i) = oneSalle Then
                Return
            End If
        Next i
        ReDim Preserve salles_(0 To UBound(salles_) + 1) As String
    Else
        ReDim Preserve salles_(0 To 0) As String
    End If
    salles_(UBound(salles_)) = oneSalle
End Sub

Public Sub EnleveSalle(oneSalle As String)
    Dim CombienSalles As Long
    If (Not Not salles_) Then
        CombienSalles = UBound(salles_)
        For i = LBound(salles_) To UBound(salles_)
            If salles_(i) = oneSalle Then
                salles_(i) = salles_(CombienSalles)
                CombienSalles = CombienSalles - 1
            End If
        Next i
        ReDim Preserve salles_(0 To CombienSalles) As String
    End If
End Sub
Public Sub AjouteEnseignant(oneEnseignant As String)
    If (Not Not enseignant_) Then
        For i = LBound(enseignant_) To UBound(enseignant_)
            If enseignant_(i) = oneEnseignant Then
                Return
            End If
        Next i
        ReDim Preserve enseignant_(0 To UBound(enseignant_) + 1) As String
    Else
        ReDim Preserve enseignant_(0 To 0) As String
    End If
    enseignant_(UBound(enseignant_)) = oneEnseignant
End Sub
Public Sub EnleveEnseignant(oneEnseignant As String)
    Dim CombienEnseignants As Long
    If (Not Not enseignant_) Then
        CombienEnseignants = UBound(enseignant_)
        For i = LBound(enseignant_) To UBound(enseignant_)
            If enseignant_(i) = oneEnseignant Then
                enseignant_(i) = enseignant_(CombienEnseignants)
                CombienEnseignants = CombienEnseignants - 1
            End If
        Next i
        ReDim Preserve enseignant_(0 To CombienEnseignants) As String
    End If
End Sub
Public Function WriteSalles() As String
    Dim ret As String
    ret = ""
    If (Not Not salles_) Then
        For i = LBound(salles_) To UBound(salles_)
            If i <> LBound(salles_) Then
                ret = ret & ", "
            End If
            ret = ret & salles_(i)
        Next i
    End If
    WriteSalles = ret
End Function
Public Function WriteStr() As String
    Dim ret As String
    ret = ue_ & " # " & Discipline & Chr(13) & Chr(10)
    ret = ret & WriteSalles()
    ret = ret & Chr(13) & Chr(10)
    If (Not Not enseignant_) Then
        For i = LBound(enseignant_) To UBound(enseignant_)
            If i <> LBound(enseignant_) Then
                ret = ret & ", "
            End If
            ret = ret & enseignant_(i)
        Next i
    End If
    If commentaires_ <> "" Then
        ret = ret & Chr(13) & Chr(10) & commentaires_
    End If
    WriteStr = ret
End Function

Public Sub Lire(cell As Range, Optional DayDate As Date = 0)
    inputStr = cell.Value
    If inputStr = "" Then
        Valid_ = False
        Exit Sub
    End If
    Set Attached_Worksheet = cell.Worksheet
    DateHourOfBegin = DayDate + cell.Offset(0, 2 - cell.Column).Value
    DateHourOfEnd = DayDate + cell.Offset(0, 2 - cell.Column).Offset(cell.MergeArea.Rows.Count, 0).Value
    TimeLength = NbHours(cell)
    Dim Tx() As String
    Dim FirstLine() As String
    Dim SecondLine() As String
    Tx = Split(inputStr, Chr(10))
    Dim CombienLigne As Integer
    Dim Loc As Integer
    CombienLigne = UBound(Tx)
    ' A B _ D _ F
    ' A B
    Loc = LBound(Tx)
    For t = LBound(Tx) To UBound(Tx)
        If Trim(Tx(t)) = "" Then
            CombienLigne = CombienLigne - 1
        Else
            If Loc <= UBound(Tx) Then
                Tx(Loc) = Tx(t)
                Loc = Loc + 1
            End If
        End If
    Next t
    ReDim Preserve Tx(0 To CombienLigne) As String

    FirstLine = Split(Tx(0), "#")
    With Me
        If UBound(FirstLine) - LBound(FirstLine) = 1 Then
            .UE = Trim(FirstLine(0))
            .Discipline = Trim(FirstLine(1))
        End If
        If (UBound(Tx) >= 1) Then
            SecondLine = Split(Tx(1), ",")
            For i = LBound(SecondLine) To UBound(SecondLine)
                .AjouteSalle (Trim(SecondLine(i)))
            Next i
            If (UBound(Tx) >= 2) Then
                Dim LstEnseignant() As String
                LstEnseignant = Split(Trim(Tx(2)), ",")
                For e = LBound(LstEnseignant) To UBound(LstEnseignant)
                    .AjouteEnseignant Trim(LstEnseignant(e))
                Next e
                If (UBound(Tx) >= 3) Then
                    For t = 3 To UBound(Tx)
                        .Commentaire = .Commentaire & Chr(13) & Tx(t)
                    Next t
                End If
            End If
        End If
    End With
    Valid_ = True
End Sub

Public Function SalleReservee(ThisRoom As String) As Boolean
    If (Not Not salles_) Then
        For k = LBound(salles_) To UBound(salles_)
            If ThisRoom = salles_(k) Then
                SalleReservee = True
                Exit Function
            End If
        Next k
    End If
    SalleReservee = False
End Function

Public Function Enseignants(ThisEnseign As String) As Boolean
    If (Not Not enseignant_) Then
        For k = LBound(enseignant_) To UBound(enseignant_)
            If ThisEnseign = enseignant_(k) Then
                Enseignants = True
                Exit Function
            End If
        Next k
    End If
    Enseignants = False
End Function

Public Function GetEnseignant(idx As Integer) As String
    If (Not Not enseignant_) Then
        If idx >= LBound(enseignant_) And idx <= UBound(enseignant_) Then
            GetEnseignant = enseignant_(idx)
        Else
            GetEnseignant = ""
        End If
    Else
        GetEnseignant = ""
    End If
End Function

Public Sub VideSalles()
    Erase salles_
End Sub
Public Sub VideEnseignants()
    Erase enseignant_
End Sub

Public Sub Reset()
    VideEnseignants
    discipline_ = ""
    ue_ = ""
    VideSalles
    commentaires_ = ""
    Valid_ = False
End Sub

Private Sub Class_Initialize()

End Sub




