Function Contains(objCollection As Object, strName As String) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = objCollection(strName)
    Contains = (Err.Number = 0)
 End Function
Sub RefreshCal()

' Erase Calendar Sheet if it exists
If Contains(Sheets, "Calendar") Then
    Application.DisplayAlerts = False
    Sheets("Calendar").Delete
    Application.DisplayAlerts = True
End If
' Duplicate Hidden Template tab as Calendar, hide again
Call DupeTab
' Add Events
Call AddEvents
End Sub
Sub DupeTab()
'
' DupeTab
'
    Sheets("Sheet1 (2)").Select
    Sheets("Template").Visible = True
    Range("B43").Select
    Sheets("Template").Select
    Sheets("Template").Copy Before:=Sheets(2)
    Sheets("Template (2)").Select
    Sheets("Template (2)").Name = "Calendar"
    Sheets("Template").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Calendar").Select
End Sub
Private Static Sub AddEvents()
'
' Copy Name from Name column Events Into Calendar

Dim col1
Sheets("Events").Select
'This loop runs until there is nothing in the Previous column
Range("A5").Select
    Do
    ' Copy Name from A Column
    Selection.Copy
    ' Store cell color as col1
    col1 = ActiveCell.Interior.Color
    ' Select the Date column, store to Date1
    ActiveCell.Offset(0, 11).Select
    Date1 = ActiveCell.Value
    ' Select Duration column
    ActiveCell.Offset(0, 1).Select
    ' Set Duration to Duration value or 1
    If ActiveCell.Value > 1 Then
    Duration = ActiveCell.Value
    Else: Duration = 1
    End If
    ' Paste A Column into Calendar
    Do
        Sheets("Calendar").Select
        ' ConvertDate selects the cell with that date in it
        Call ConvertDate(Date1)
        ' Go down to next empty cell.
        Do
            ActiveCell.Offset(1, 0).Select
        Loop Until IsEmpty(ActiveCell.Offset(0, 0))
        ' Paste text only:
        Selection.PasteSpecial Paste:=xlPasteValues
        ' ActiveSheet.Paste
        ActiveCell.Interior.Color = col1
        Duration = Duration - 1
        Date1 = Date1 + 1
    Loop Until Duration = 0
    ' Return to Events and move down a cell
    Sheets("Events").Select
    ActiveCell.Offset(1, -12).Select
    Loop Until IsEmpty(ActiveCell.Offset(0, 0))
    ' Show Calendar
    Sheets("Calendar").Select
    Application.Goto Range("$A$1")
End Sub
Private Static Sub ConvertDate(Date1)
' Convert Dates to row/column in Calendar
Dim FindString As String
FindString = Day(Date1)
'November 2014
If Month(Date1) = 11 And Year(Date1) = 2014 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B5:H40")
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
'December 2014
ElseIf Month(Date1) = 12 And Year(Date1) = 2014 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B51:H79")
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
'January 2015
ElseIf Month(Date1) = 1 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B98:H126")
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
'February 2015
ElseIf Month(Date1) = 2 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B145:H166")
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
'March 2015
ElseIf Month(Date1) = 3 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B192:H220")
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
'April 2015
ElseIf Month(Date1) = 4 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B239:H267")
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
'May 2015
ElseIf Month(Date1) = 5 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B286:H321")
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
'June 2015
ElseIf Month(Date1) = 6 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B333:H361")
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
'July 2015
ElseIf Month(Date1) = 7 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B380:H408")
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
'August 2015
ElseIf Month(Date1) = 8 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B427:H462")
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
'September 2015
ElseIf Month(Date1) = 9 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B474:H502")
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
'October 2015
ElseIf Month(Date1) = 10 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B521:H549")
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
'November 2015
ElseIf Month(Date1) = 11 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B568:H596")
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
'December 2015
ElseIf Month(Date1) = 12 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B615:H643")
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

Else
    MsgBox ("Error: " & Date1 & " is not yet in calendar sheet.")
    Application.Goto Range("$B$660")
    ActiveCell.FormulaR1C1 = "Errors:"
End If
End Sub


