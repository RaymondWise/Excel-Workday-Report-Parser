Attribute VB_Name = "WorkdayLineFeedMgr"
Option Explicit
'==========================================
'MIT License
'Copyright (c) <2016> <Raymond Wise> <https://github.com/RaymondWise/Excel-Workday-Report-Parser> @raymondwise
'==========================================

'#Procedures for removing and parsing Workday reports with LineFeed characters in column

Public Sub ParseIntoColumns()
'Parse column to the right (text to columns)
    Dim confirmOverwrite As String
    confirmOverwrite = MsgBox("Do you want to overwrite all data to the right of your selection?", vbYesNo)
    If confirmOverwrite = vbNo Then Exit Sub

    Dim lastRow As Long
    lastRow = 1

    Dim workingRange As Range
    Set workingRange = UserSelectRange(lastRow)
    If workingRange Is Nothing Then Exit Sub
    Dim workingSheet As Worksheet
    Set workingSheet = workingRange.Parent
    Dim workingColumn As Long
    workingColumn = workingRange.Column

    Application.ScreenUpdating = False
    workingRange.TextToColumns _
    Destination:=workingRange, _
    DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierNone, _
        ConsecutiveDelimiter:=True, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=True, OtherChar:=vbLf

    With workingSheet.UsedRange
        .WrapText = False
        .EntireColumn.AutoFit
    End With
    
    Application.ScreenUpdating = True
End Sub


Public Sub ParseIntoRows()
'Parse column downward, inserting rows
    Dim lastRow As Long
    lastRow = 1

    Dim workingRange As Range
    Set workingRange = UserSelectRange(lastRow)

    If workingRange Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Dim workingColumn As Long
    workingColumn = workingRange.Column
    Dim currentRow As Long
    Dim cellToParse As Range
    Dim stringParts() As String

    For currentRow = lastRow To 2 Step -1

        Set cellToParse = Cells(currentRow, workingColumn)
        stringParts = Split(cellToParse, vbLf)
        If Len(Join(stringParts)) > 0 Then CreateNewRows stringParts(), cellToParse
    
    Next currentRow

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub


Private Sub CreateNewRows(ByRef partsOfString() As String, ByVal targetCellToParse As Range)
    targetCellToParse.Value = partsOfString(0)
    Dim i As Long
    For i = 1 To UBound(partsOfString)
        If Len(partsOfString(i)) > 0 Then
            targetCellToParse.EntireRow.Copy
            targetCellToParse.EntireRow.Insert shift:=xlDown
            targetCellToParse.Offset(-1) = partsOfString(i)
        End If
    Next i
End Sub


Private Function UserSelectRange(ByRef lastRow As Long) As Range
    Set UserSelectRange = Nothing
    Dim columnToParse As Range

    Set columnToParse = GetUserInputRange
    If columnToParse Is Nothing Then Exit Function

    If columnToParse.Columns.Count > 1 Then
        MsgBox "You selected multiple columns. Exiting.."
        Exit Function
    End If

    Dim columnLetter As String
    columnLetter = ColumnNumberToLetter(columnToParse)

    Dim result As String
    result = MsgBox("The column you've selected to parse is column " & columnLetter, vbOKCancel)
    If result = vbCancel Then
        MsgBox "Process Cancelled."
    Exit Function
    End If

    lastRow = Cells(Rows.Count, columnToParse.Column).End(xlUp).Row
    Set UserSelectRange = Range(Cells(2, columnToParse.Column), Cells(lastRow, columnToParse.Column))

End Function


Private Function GetUserInputRange() As Range
    'This is segregated because of how excel handles cancelling a range input
    Dim userAnswer As Range
    On Error GoTo InputError
    Set userAnswer = Application.InputBox("Please click a cell in the column to parse", "Column Parser", Type:=8)
    Set GetUserInputRange = userAnswer
    Exit Function
InputError:
    Set GetUserInputRange = Nothing
End Function


Private Function ColumnNumberToLetter(ByVal selectedRange As Range) As String
    'Convert column number to column letter representation

    Dim rowBeginningPosition As Long
    rowBeginningPosition = InStr(2, selectedRange.Address, "$")
    Dim columnLetter As String
    columnLetter = Mid(selectedRange.Address, 2, rowBeginningPosition - 2)
    
    'Handles if the user selected an entire column
    If Right(columnLetter, 1) = ":" Then
        ColumnNumberToLetter = Left(columnLetter, Len(columnLetter) - 1)
    Else: ColumnNumberToLetter = columnLetter
    End If
    
End Function











