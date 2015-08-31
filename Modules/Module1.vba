' This following parts are made to be able to use timer.
#If VBA7 Then
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
"QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias _
"QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
Private Declare Function getFrequency Lib "kernel32" Alias _
"QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function getTickCount Lib "kernel32" Alias _
"QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If

Function NbHours(rCell As Range) As Double
' Returns true if referenced cell is Merged
    Dim i As Long
    Dim res As Double
    res = 0
    i = rCell.MergeArea.Cells(1, 1).Row
    Do While i < rCell.MergeArea.Cells(1, 1).Row + rCell.MergeArea.Rows.Count
        res = res + Worksheets("Annuel").Range("B" & i + 1).Value - Worksheets("Annuel").Range("B" & i).Value
        i = i + 2
    Loop
    NbHours = res
End Function

Function GetDateFromCell(rCell As Range) As Date
    r = rCell.Cells(1, 3 - rCell.Column).End(xlDown).End(xlUp).Row - 1 - rCell.Row
    
    GetDateFromCell = CDate(rCell.Cells(r, 1).MergeArea.Cells(1, 1))
End Function
Public Function MicroTimer() As Double
'
' returns seconds
' uses Windows API calls to the high resolution timer
'
Dim cyTicks1 As Currency
Dim cyTicks2 As Currency
Static cyFrequency As Currency
'
MicroTimer = 0
'
' get frequency
'
If cyFrequency = 0 Then getFrequency cyFrequency
'
' get ticks
'
getTickCount cyTicks1
getTickCount cyTicks2
If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
'
' calc seconds
'
If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
End Function
Function CountHour(Enseignant As String, Discipline As String, UE As String, Where As Range) As Double
    Dim c As Range
    Dim ch As Integer
    CountHour = 0
    For Each c In Where
        Dim a As String, ca As String
        ca = c.Address
        a = c.MergeArea.Cells(1, 1).Address
        If (a = ca) Then
            Dim MyCren As New Creneau
            MyCren.Reset
            MyCren.Lire c
            If (( _
                        Enseignant = "" _
                    Or MyCren.Enseignants(Enseignant) _
                    ) _
                And ( _
                        Discipline = "" _
                    Or MyCren.Discipline = Discipline _
                    ) _
                And ( _
                        UE = "" _
                    Or MyCren.UE = UE _
                    ) _
                ) Then
                CountHour = CountHour + MyCren.TimeDelta
            End If
        End If
    Next c
End Function



