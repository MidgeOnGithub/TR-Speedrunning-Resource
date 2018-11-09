<<<<<<< HEAD
Attribute VB_Name = "mPublicFunctions"
Option Explicit

'This function is takes the user's cell in an Ammo sheet and translates it to the corresponding Archive sht row
Public Function CollapseEnemyAndLevel()

    'Code to get row number from ActiveCell here!!!

End Function

'This function is meant to take an Archive sht row and translate it to the exact cell in the Ammo sht
Function ExpandEnemyAndLevel(ElementIndex As Integer, TblKills As ListObject) As Range
    'ElementIndex comes from Row in which item appears in archive table.
    'TblKills refers to ActiveSheet's (the one user is giving input in) TblKills.
    'Output describes which cell

    Dim ColStart As Integer: ColStart = TblKills.Range.Column
    Dim ColEnd As Integer: ColEnd = TblKills.ListColumns.Count - 1
    'Subtract 1 for TR2 because HSH is exempt from weapon selection.
    
    Dim RowStart As Integer: RowStart = TblKills.Range.Row
    Dim RowEnd As Integer: RowEnd = TblKills.ListRows.Count
    Dim ArrayToCell() As Variant

    Dim i As Integer, j As Integer
    Dim Count As Integer: Count = 0
    For i = ColStart To i = ColEnd
        For j = RowStart To j = RowEnd
            If Not ActiveSheet.Range(i, j).Value = "" Then
                Count = Count + 1
                ArrayToCell(Count) = Range(i, j)
            End If
        Next
    Next
    ExpandEnemyAndLevel = ArrayToCell(ElementIndex)
    'CollapseEnemyAndLevel = Count
End Function

Public Function CollapseRun() As Long  
    
    Dim ActSht As Worksheet
    Set ActSht = ThisWorkbook.ActiveSheet
    
    'Determine which ruleset is applicable ================
    Dim intRuleset As Integer
    If Not InStr(ActSht.Name, "Any%") = 0 Then
        intRuleset = 1
    ElseIf Not InStr(ActSht.Name, "Secrets%") = 0 Then
        intRuleset = 3
    ElseIf Not InStr(ActSht.Name, "100%") = 0 Then
        intRuleset = 5
    Else  'Error checking
        CollapseRun = "ShtRenamed"
        Exit Function
    End If
    'Add one if glitchless since Ammo sheets are grouped together
    If Not InStr(ActSht.Name, "Glitchless" = 0) Then intRuleset = intRuleset + 1
    '=====================================================|
    
    'Get "start" of group of columns in archive sheet; final value to be evaluated by adding Version
    CollapseRun = (intRuleset - 1) * 4 + 1
    
    'Variables to get the range of cells to loop through when finding VersionCount.
    Dim lVisRows As Long: lVisRows = 0
    Do
        lVisRows = lVisRows + 1
    Loop Until ActSht.Rows(i).Visible = False
    lVisRows = lVisRows - 1  'Subtracting 1 to get last True value
    
    Dim lVisCols As Long: lVisCols = 0
    Do
        lVisCols = lVisCols + 1
    Loop Until ActSht.Rows(i).Visible = False
    lVisCols = lVisCols - 1
    
    'Determine which Version is applicable ================
    Dim EndCell As Range
    Set EndCell = ActSht.Cells(lVisRows, lVisCols)  'Bottom-right-most cell
    Dim VisibleRange As Range
    Set VisibleRange = ActSht.Range(ActSht.Cells(1, 1), EndCell)  'Cell Range to loop through.

    Dim intCzCells As Integer: intCzCells = 0
    Dim Cell As Range  'Iterates through VisibleRange
    Dim arrCzValues As Variant  'Stores CzCell values
    For Each Cell In VisibleRange
        On Error GoTo NoName  'Workaround for .Name.Name throwing errors if Cell has not been manually renamed
        If InStr(Cell.Name.Name, "CheckCell") <> 0 Then
            intCzCells = intCzCells + 1
            ReDim Preserve arrCz(intCzCells - 1)  'Array indexed from 0.
            arrCz(intCzCells - 1) = Cell.Address  'Subtract 1 since arrays index from 0
            arrCzValues(intCzCells - 1) = Cell.Value
        End If
    NoName:  'If error went straight to NextIteration without Resume, the code would run in "debug" mode
            Resume NextIteration
    NextIteration:
    Next Cell

    Dim intCzCount As Integer
    intCzCount = intCzCells ^ 2  'Each CzCell is a Boolean dropdown, this returns combinations which must be sorted in "binary" fashion in archive sheets
    Dim Version As Integer
    Dim i As Variant
    For Each i In arrCzValues
        If arrCzValues(i) = "Yes" Then Version = Version + 2 ^ (i - 1)
    Next i
    '====================================================||
    CollapseRun = CollapseRun + (Version - 1)
End Function

Function ExpandRun(Output As String, ElementIndex As Integer, VersionCount As Integer) As Integer 'Returns ???
    Select Case Output
        Case "Ruleset"
            ExpandRun = ElementIndex Mod VersionCount
        Case "Version"
            ExpandRun = Application.WorksheetFunction.RoundUp((ElementIndex / VersionCount), 0)
        Case Else
            MsgBox "Code error: incorrect Parameter passed to ExpandRun. Process Terminated."
    End Select
End Function
=======
Attribute VB_Name = "mPublicFunctions"
Option Explicit

'This function takes the user's cell in an Ammo sheet and translates it to the corresponding Archive sht row
Public Function CollapseEnemyAndLevel()

    Dim Sht As Worksheet
    Set Sht = ThisWorkbook.ActiveSheet
    
    

End Function

'This function is meant to take an Archive sht row and translate it to the exact cell in the Ammo sht
Public Function ExpandEnemyAndLevel(ElementIndex As Integer, TblKills As ListObject) As Range
    'ElementIndex comes from Row in which item appears in archive table.
    'TblKills refers to ActiveSheet's (the one user is giving input in) TblKills.
    'Output describes which cell

    Dim ColStart As Integer: ColStart = TblKills.Range.Column
    Dim ColEnd As Integer: ColEnd = TblKills.ListColumns.Count - 1
    'Subtract 1 for TR2 because HSH is exempt from weapon selection.
    
    Dim RowStart As Integer: RowStart = TblKills.Range.Row
    Dim RowEnd As Integer: RowEnd = TblKills.ListRows.Count
    
    Dim InputCells() As Variant
    Dim i As Integer, j As Integer
    Dim Count As Integer: Count = 0
    For i = ColStart To i = ColEnd
        For j = RowStart To j = RowEnd
            If Not ActiveSheet.Range(i, j).Value = "" Then
                Count = Count + 1
                InputCells(Count) = Range(i, j)
            End If
        Next
    Next

    ExpandEnemyAndLevel = InputCells(ElementIndex)
End Function

Public Function CollapseRun() As Long
    
    Dim ActSht As Worksheet
    Set ActSht = ThisWorkbook.ActiveSheet
    
    'Determine which ruleset is applicable ================
    Dim intRuleset As Integer
    If Not InStr(ActSht.Name, "Any%") = 0 Then
        intRuleset = 1
    ElseIf Not InStr(ActSht.Name, "Secrets%") = 0 Then
        intRuleset = 3
    ElseIf Not InStr(ActSht.Name, "100%") = 0 Then
        intRuleset = 5
    Else  'Error checking
        CollapseRun = "ShtRenamed"
        Exit Function
    End If
    'Add one if glitchless since Ammo sheets are grouped together
    If Not InStr(ActSht.Name, "Glitchless" = 0) Then intRuleset = intRuleset + 1
    '=====================================================|
    
    'Get "start" of group of columns in archive sheet; final value to be evaluated by adding Version
    CollapseRun = (intRuleset - 1) * 4 + 1
    
    'Variables to get the range of cells to loop through when finding VersionCount.
    Dim lVisRows As Long: lVisRows = 0
    Do
        lVisRows = lVisRows + 1
    Loop Until ActSht.Rows(i).Visible = False
    lVisRows = lVisRows - 1  'Subtracting 1 to get last True value
    
    Dim lVisCols As Long: lVisCols = 0
    Do
        lVisCols = lVisCols + 1
    Loop Until ActSht.Rows(i).Visible = False
    lVisCols = lVisCols - 1
    
    'Determine which Version is applicable ================
    Dim EndCell As Range
    Set EndCell = ActSht.Cells(lVisRows, lVisCols)  'Bottom-right-most cell
    Dim VisibleRange As Range
    Set VisibleRange = ActSht.Range(ActSht.Cells(1, 1), EndCell)  'Cell Range to loop through.

    Dim intCzCells As Integer: intCzCells = 0
    Dim Cell As Range  'Iterates through VisibleRange
    Dim arrCzValues As Variant  'Stores CzCell values
    For Each Cell In VisibleRange
        On Error GoTo NoName  'Workaround for .Name.Name throwing errors if Cell has not been manually renamed
        If InStr(Cell.Name.Name, "CheckCell") <> 0 Then
            intCzCells = intCzCells + 1
            ReDim Preserve arrCz(intCzCells - 1)  'Array indexed from 0.
            arrCz(intCzCells - 1) = Cell.Address  'Subtract 1 since arrays index from 0
            arrCzValues(intCzCells - 1) = Cell.Value
        End If
    NoName:  'If error went straight to NextIteration without Resume, the code would run in "debug" mode
            Resume NextIteration
    NextIteration:
    Next Cell

    Dim intCzCount As Integer
    intCzCount = intCzCells ^ 2  'Each CzCell is a Boolean dropdown, this returns combinations which must be sorted in "binary" fashion in archive sheets
    Dim Version As Integer
    Dim i As Variant
    For Each i In arrCzValues
        If arrCzValues(i) = "Yes" Then Version = Version + 2 ^ (i - 1)
    Next i
    '====================================================||
    CollapseRun = CollapseRun + (Version - 1)
End Function

Public Function ExpandRun(Output As String, ElementIndex As Integer, VersionCount As Integer) As Integer  'Returns ???
    Select Case Output
        Case "Ruleset"
            ExpandRun = ElementIndex Mod VersionCount
        Case "Version"
            ExpandRun = WorksheetFunction.RoundUp((ElementIndex / VersionCount), 0)
        Case Else
            MsgBox "Code error: incorrect Parameter passed to ExpandRun. Process Terminated."
    End Select
End Function
>>>>>>> d7b9fa5... mWeaponsAvailable fixed, extra spaces removed
