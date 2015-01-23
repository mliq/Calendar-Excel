Dim rng2 As Range
Dim row1 As Integer
Dim row2 As Integer
row1 = 5
row2 = 40
Set rng2 = Range("B" & row1 & ":" & "H" & row2)

'November 2014
If Month(Date1) = 11 And Year(Date1) = 2014 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range(rng2)
                Set rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not rng Is Nothing Then
                    Application.Goto rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If