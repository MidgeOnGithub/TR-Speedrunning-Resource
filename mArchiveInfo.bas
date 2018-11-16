Attribute VB_Name = "mArchiveInfo"
Option Explicit

'This function takes the user's cell in an Ammo sheet and translates it to the corresponding Archive sht row
Function CollapseEnemyAndLevel(ItemRge As Range) As Long
    
    'Shorthand
    Dim Sht As Worksheet
    Set Sht = ThisWorkbook.ActiveSheet
    
    'Determine which ruleset is applicable ================
    Dim strRun As String
    strRun = FindRunType(Sht)
    
    Dim tblKills As ListObject
    Set tblKills = Sht.ListObjects("tbl" + strRun + "Kills")
    
    'Getting column and row data to bound loop search
    Dim ColStart As Integer: ColStart = tblKills.Range.Column
    Dim ColEnd As Integer: ColEnd = ColStart + (tblKills.ListColumns.Count - 2)
    'Subtract an extra 1 for TR2 because HSH is exempt from weapon selection.
    Dim RowStart As Integer: RowStart = tblKills.Range.Row
    Dim RowEnd As Integer: RowEnd = RowStart + (tblKills.ListRows.Count - 1)
    
    Dim i As Integer, j As Integer
    Dim Count As Integer: Count = 0
    For j = ColStart To ColEnd
        
        For i = RowStart To RowEnd
            
            If Not Sht.Cells(i, j).Value = "" Then
                Count = Count + 1
                
                If i = ItemRge.Row And j = ItemRge.Column Then
                    CollapseEnemyAndLevel = Count
                    Exit Function
                End If
            
            End If
        Next i
    Next j
    
    CollapseEnemyAndLevel = -403
    
End Function

'This function is meant to take an Archive sht row and translate it to the exact cell in the Ammo sht
Function ExpandEnemyAndLevel(ElementIndex As Integer, tblKills As ListObject) As Range
    'ElementIndex comes from Row in which item appears in archive table.
    'tblKills refers to ActiveSheet's (the one user is giving input in) tblKills.
    
    'Getting column and row data to bound loop search
    Dim ColStart As Integer: ColStart = tblKills.Range.Column
    Dim ColEnd As Integer: ColEnd = tblKills.ListColumns.Count - 1
    'Subtract 1 for TR2 because HSH is exempt from weapon selection.
    Dim RowStart As Integer: RowStart = tblKills.Range.Row
    Dim RowEnd As Integer: RowEnd = tblKills.ListRows.Count
    
    Dim InputCells() As Variant
    Dim i As Integer, j As Integer
    Dim Count As Integer: Count = 0
    For i = ColStart To ColEnd
        
        For j = RowStart To RowEnd
            
            If Not ActiveSheet.Range(i, j).Value = "" Then
                Count = Count + 1
                InputCells(Count) = Range(i, j).Address
            End If
        
        Next j
    Next i
    
    'Output describes which cell to use
    ExpandEnemyAndLevel = InputCells(ElementIndex)
End Function

Function CollapseRun() As Long
    
    'Shorthand
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
        CollapseRun = -404
        Exit Function
    End If
    'Add one if glitchless since Ammo sheets are grouped together
    If Not InStr(ActSht.Name, "Glitchless") = 0 Then intRuleset = intRuleset + 1
    '=====================================================|
    
    'Get "start" of group of columns in archive sheet; final value to be evaluated by adding Version
    CollapseRun = (intRuleset - 1) * 4 + 1
    
    'Variables to get the range of cells to loop through when finding VersionCount.
    Dim lVisRows As Long: lVisRows = 0
    Do
        lVisRows = lVisRows + 1
    Loop Until ActSht.Rows(lVisRows).Hidden
    lVisRows = lVisRows - 1  'Subtracting 1 to get last True value
    
    Dim lVisCols As Long: lVisCols = 0
    Do
        lVisCols = lVisCols + 1
    Loop Until ActSht.Rows(lVisCols).Hidden
    lVisCols = lVisCols - 1
    
    'Determine which Version is applicable ================
    Dim EndCell As Range
    Set EndCell = ActSht.Cells(lVisRows, lVisCols)  'Bottom-right-most cell
    Dim VisibleRange As Range
    Set VisibleRange = ActSht.Range(ActSht.Cells(1, 1), EndCell)  'Cell Range to loop through.

    Dim intCzCells As Integer: intCzCells = 0
    Dim Cell As Range  'Iterates through VisibleRange
    Dim arrCzValues() As Variant  'Stores CzCell values
    For Each Cell In VisibleRange
        On Error GoTo NoName  'Workaround for .Name.Name throwing errors if Cell has not been manually renamed
        If InStr(Cell.Name.Name, "CheckCell") <> 0 Then
            intCzCells = intCzCells + 1
            ReDim Preserve arrCzValues(1 To intCzCells)
            arrCzValues(intCzCells) = Cell.Value
        End If
    NoName:  'If error went straight to NextIteration without Resume, the code would run in "debug" mode
            Resume NextIteration
    NextIteration:
    Next Cell

    Dim intCzCount As Integer
    intCzCount = intCzCells ^ 2  'Each CzCell is a Boolean dropdown, this returns combinations which must be sorted in "binary" fashion in archive sheets
    Dim Version As Integer: Version = 0
    Dim i As Integer
    For i = 1 To (UBound(arrCzValues) - LBound(arrCzValues) + 1)
        If arrCzValues(i) = "Yes" Then Version = Version + 2 ^ (i - 1)
    Next i
    '====================================================||
    CollapseRun = CollapseRun + (Version - 1)
    CollapseRun = CollapseRun + 1
End Function

Function ExpandRun(Output As String, ElementIndex As Integer, VersionCount As Integer) As Integer  'Returns ???
    Select Case Output
        Case "Ruleset"
            ExpandRun = ElementIndex Mod VersionCount
        Case "Version"
            ExpandRun = WorksheetFunction.RoundUp((ElementIndex / VersionCount), 0)
        Case Else
            MsgBox "Code error: incorrect Parameter passed to ExpandRun. Process Terminated."
    End Select
End Function

'NVIndirect stands for non-volatile INDIRECT function, it's a workaround for Excel's INDIRECT function being volatile
'Volatility means the function will recalculate with *any* change in the workbook, which is terrible for functions relying on ActiveSheet
'This makes the functions static, which is also a problem, but can be fixed
Function NVIndirect(Eval As String) As Variant
    NVIndirect = Range(Eval)
End Function

'Taken from https://stackoverflow.com/a/15366979/10466817
'Turns a number meant for a row or column into a letter
'Useful in combination with NVIndirect as an alternative to INDIRECT(ADDRESS) as both INDIRECT and ADDRESS are volatile
Function ToLetter(ColNum As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColNum
    
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    
    ToLetter = s
End Function

'This is standalone from main for use by tblAmmo cells
Function WeaponName(Rge As Range) As String
    
    Dim Sht As Worksheet  'Shorthand
    Set Sht = Rge.Worksheet
    
    'Determine which ruleset is applicable ================
    Dim strRun As String
    strRun = FindRunType(Sht)
    
    Dim tblAmmo As ListObject
    Set tblAmmo = Sht.ListObjects("tbl" & strRun & "Ammo")

    Dim RowStart As Integer: RowStart = tblAmmo.Range.Row
    Dim RowEnd As Integer: RowEnd = RowStart + (tblAmmo.ListRows.Count - 1)

    Dim Weapon As Collection: Set Weapon = PopulateColl("Weapon")
    Dim WeaponIndex As Integer

    Dim Row As Long
    For Row = RowStart To RowEnd
            
        If Row = Rge.Row Then
            WeaponIndex = (Row + 1) - RowStart
            Exit For
        End If
        
    Next Row
    
    WeaponName = Weapon("Weapon" & WeaponIndex)
    
End Function
