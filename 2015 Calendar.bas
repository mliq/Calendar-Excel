Attribute VB_Name = "Module1"
Option Explicit
Private Errs As Integer
Private objTemplate As Excel.Worksheet
Private objEvents As Excel.Worksheet
Private objCalendar As Excel.Worksheet
Private xlApp As Excel.Application
Private xlBook As Excel.Workbook
Private xlSheet As Excel.Worksheet
' Private hideRow As Integer

Function Contains(objCollection As Object, strName As String) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = objCollection(strName)
    Contains = (Err.Number = 0)
 End Function
Function testVis(testCell As Range) As Integer

' Return 0 if Cell is invisible, 1 if visible

objEvents.Select

Dim isect As Range
Dim visCells As Range

Set visCells = ActiveSheet.Range("A1:A200").Rows.SpecialCells(xlCellTypeVisible)
Set isect = Intersect(testCell, visCells)

If Not (isect Is Nothing) Then
  '   MsgBox ("Cell is visible")
    testVis = 1
    Exit Function
End If
' MsgBox (isect)
testVis = 0
End Function
Private Sub ResetWindowView()
 
' store current worksheet, for later reference
Dim thisSht As Excel.Worksheet
Set thisSht = ActiveSheet
 
Dim sht As Excel.Worksheet
Dim currentWindow As Window
 
' loop through each worksheet and set view
For Each sht In Excel.Worksheets
  If TypeName(sht) = "Worksheet" Then
    sht.Activate
    Set currentWindow = ActiveWindow
    currentWindow.View = xlNormalView
  End If
Next sht
 
' go back to original sheet
thisSht.Activate
 
End Sub
Public Sub UserInput()
Dim HideRow As Long
Dim message, title, defaultValue As String

message = "Enter the Row Number of the Blue Title row of the first month you wish to remain visible:"
' title = "Hide Passed Months (Default 96 begins at January 2015)"
title = "Hide Months"
defaultValue = 96
HideRow = Application.InputBox(message, title, defaultValue, Type:=1)
If HideRow <= 95 Then
    Exit Sub
Else
    Call RefreshCal(HideRow)
End If
End Sub
Sub RefreshCal(Optional HideRow As Long)

If HideRow = 0 Then
    HideRow = Range("P3").Value
End If

Set xlApp = Application
Set xlBook = ThisWorkbook ' xlApp.Workbooks.Add
Set objTemplate = xlApp.Sheets("Template")
Set objEvents = xlApp.Sheets("Events")
' Hide screen
Application.ScreenUpdating = False
' Set up Error Counter
Errs = 0
' Erase Calendar Sheet if it exists
If Contains(Sheets, "Calendar") Then
    Application.DisplayAlerts = False
    Sheets("Calendar").Delete
    Application.DisplayAlerts = True
End If

' Duplicate Hidden Template tab as Calendar, hide again, passing in Hiderow variable inputted, or 0 otherwise

' Call DupeTab
Call DupeTab
Set objCalendar = Application.Sheets("Calendar")

' Add Events
' Hide 2014 Nov and Dec Rows by Default...
If HideRow = 0 Then
Call AddEvents(96) 'Rows("3:94").Hidden = True
Else
Call AddEvents(HideRow)
End If

' Save HideRow Variable for future:
Range("P3").Value = HideRow


Application.ScreenUpdating = True
'    Dim strPrompt As String
 '   Dim iret As Integer
  '  Dim strTitle As String
    
   ' strPrompt = "Errors number " & Errs
    'strTitle = "Title"
    'iret = MsgBox(strPrompt, vbOKOnly)


End Sub
Private Sub DupeTab()
'
' DupeTab
'
' Dim oXL As Excel.Application
' Set oXL = Application
' oXL.Visible = True

Call ResetWindowView
'Here's the actual sheet addition code
With Application
.ScreenUpdating = False
.EnableEvents = False
.DisplayAlerts = False
End With
'Erase old
'Worksheets(“Calendar”).Delete

'Add and name the new sheet
Worksheets.Add
With ActiveSheet
.Name = "Calendar"
.Move After:=Worksheets(1)
ActiveSheet.PageSetup.Orientation = xlLandscape
End With

'Make the Template sheet visible, and copy it
With Worksheets("Template")
.Visible = xlSheetVisible
.Activate
End With
' ActiveWindow.View = xlNormalView
Cells.Copy
'Re-activate the new worksheet, and paste
Worksheets("Calendar").Activate
Cells.Select
ActiveSheet.Paste
Rows("1:800").RowHeight = 11.16

With Application
.CutCopyMode = False
.Goto Range("A1"), True
End With

With Application
' .ScreenUpdating = True
.EnableEvents = True
.DisplayAlerts = True
End With

'Make Template Sheet invisible again
With Worksheets("Template")
.Visible = xlSheetHidden
End With

End Sub
Private Sub AddEvents(HideRow As Long)
Dim col1
Dim val
Dim Date1
Dim Duration
' Dim xlCell As Range
' Dim total As Integer
' total = 0
Dim c As Range
Dim d As Range
' Dim ErrStart As Integer

objEvents.Activate

' Cycle through each row that has an A column entry
For Each c In Range("$A$4:$A$" & Cells(Rows.Count, "A").End(xlUp).Row)
    ' ErrStart = Errs
    ' add visibility filter
    If testVis(c) = 1 Then
  '      total = total + 1 ' do stuff
        val = c.Value
        col1 = c.Interior.Color
        c.Offset(0, 11).Select
        Date1 = ActiveCell.Value
        ' Select Duration column
        c.Offset(0, 12).Select
        ' Set Duration to Duration value or 1
        If ActiveCell.Value > 1 Then
            Duration = ActiveCell.Value
            Else: Duration = 1
        End If
        
        ' Put c into Calendar
        Do
            ' ConvertDate selects the cell with that date in it
            Call ConvertDate(Date1)
               ' Go down vertically to next empty cell in that date
                Do
                    ActiveCell.Offset(1, 0).Select
                Loop Until IsEmpty(ActiveCell.Offset(0, 0))
                ' Paste text only:
                ActiveCell.Value = val
                ActiveCell.Interior.Color = col1
                Duration = Duration - 1
                Date1 = Date1 + 1
            Loop Until Duration = 0
        ' MsgBox (val & col1 & Date1 & Duration)
    End If
    ' If ErrStart < Errs Then
        ' With c.Offset(0, 11)
        ' .Interior.ColorIndex = 3
    ' End If
Next c
' MsgBox (total)
    ' Show Calendar
    objCalendar.Select
    Application.Goto Range("$A$1"), Scroll:=True
    If Errs > 0 Then
        MsgBox ("Error: " & Errs & " dates are not yet in calendar sheet, these items have been listed on the last page under 'Errors:'")
        Application.Goto Range("$B$660"), Scroll:=True
        ActiveCell.Value = "Errors:"
    End If

With ActiveSheet
    .Rows(3 & ":" & HideRow - 2).Hidden = True
End With

End Sub
Private Static Sub ConvertDate(Date1)
' Convert Dates to row/column in Calendar
Dim FindString As String
FindString = Day(Date1)
Dim Rng As Range

'January 2015
If Month(Date1) = 1 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B98:H126")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'February 2015
ElseIf Month(Date1) = 2 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B145:H166")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'March 2015
ElseIf Month(Date1) = 3 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B192:H220")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'April 2015
ElseIf Month(Date1) = 4 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B239:H267")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'May 2015
ElseIf Month(Date1) = 5 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B286:H321")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'June 2015
ElseIf Month(Date1) = 6 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B333:H361")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'July 2015
ElseIf Month(Date1) = 7 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B380:H408")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'August 2015
ElseIf Month(Date1) = 8 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B427:H462")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'September 2015
ElseIf Month(Date1) = 9 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B474:H502")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'October 2015
ElseIf Month(Date1) = 10 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B521:H549")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'November 2015
ElseIf Month(Date1) = 11 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B568:H596")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
'December 2015
ElseIf Month(Date1) = 12 And Year(Date1) = 2015 Then
    If Trim(FindString) <> "" Then
            With Sheets("Calendar").Range("B615:H643")
                Set Rng = .Find(what:=FindString, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                Else
                    MsgBox "Nothing found"
                End If
            End With
        End If
Else
    Errs = Errs + 1
    Application.Goto Sheets("Calendar").Range("$B$660")
End If
End Sub



