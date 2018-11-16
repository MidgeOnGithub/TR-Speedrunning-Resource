Attribute VB_Name = "mRunInfo"
Option Explicit

Function FindSelectionData(Data As String, Rge As Range, tblKills As ListObject) As Integer
    'Works because each Ammo sheet's user input areas are actually Objects.
    'Uses Rge's row/column value relative to the sheet, subtracting tblKills's first row/column value
    Dim RgeProperty As Integer, TblStart As Integer
    If Data = "Enemy" Then
        RgeProperty = Rge.Row
        TblStart = tblKills.Range.Row
    ElseIf Data = "Level" Then
        RgeProperty = Rge.Column
        TblStart = tblKills.Range.Column
    Else
        MsgBox "Code error: wrong FindSelectionData argument. Process terminated."
        End
    End If
    
    FindSelectionData = RgeProperty - TblStart + 1  'Collections are indexed from 1.
End Function

Function FindRunType(Sht As Worksheet) As String
    'Check current worksheet name to determine RunType.
    If InStr(Sht.Name, "Any%") = 1 Then
        FindRunType = "Any"
    ElseIf InStr(Sht.Name, "Secrets%") = 1 Then
        FindRunType = "Secrets"
    ElseIf InStr(Sht.Name, "100%") = 1 Then
        FindRunType = "100"
    Else
        Debug.Print "RunType not found in Sheet name."
        MsgBox "RunType not found at beginning of Sheet name. Process terminated."
        End
    End If
End Function

Function IsGlitchless(Sht As Worksheet) As Boolean
    IsGlitchless = False
    'Check current worksheet name to determine RunType.
    If InStr(Sht.Name, "Glitchless") <> 0 Then
        IsGlitchless = True
    End If
End Function

Function IsNewGamePlus(Level As Collection, LevelSelected As Integer, Sht As Worksheet) As Boolean
    Dim CheckCellValue As String
    CheckCellValue = Sht.Range("NGCheckCell")
    
    Select Case CheckCellValue
        Case "Yes":
            IsNewGamePlus = True
            'Needed for levels where weapons are taken despite playing NG+.
            If Level(LevelSelected) = "Offshore Rig" Then IsNewGamePlus = False  'Hardcoded level name.
        Case "No":
            IsNewGamePlus = False
        Case Else:
            MsgBox "NGCheckCell renamed or not ""Yes""/""No"". Process Terminated."
            End
    End Select
End Function
