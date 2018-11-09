<<<<<<< HEAD
Attribute VB_Name = "mMainFunctions"
Option Explicit

Function AreYouSure()
    If MsgBox("Do you wish to input stated kill values for Enemy in LevelName?", vbInformation + vbYesNo, "Are You Sure?") = vbNo Then
        '!!! Something goes here to deal with "no" response
    End If
End Function

Function FindItemIndex(Coll As Collection, Key As String) As Integer
    Dim i As Integer 'Loops over Coll to determine Key's repective index. VBA has no native procedure for this.
    For i = 1 To Coll.Count
        If Coll(i) = Coll(Key) Then   
            Exit For  'Found it
        End If
    Next
    
    FindItemIndex = i
    
    If Coll(FindItemIndex) <> Coll(Key) Then  'Check if all items were processed but no match key exists for last value
        MsgBox "Coll(FindItemIndex) failed to = Coll(Key), process terminated."
        End
    End If
End Function

Function FindSelectionData(Data As String, Rge As Range, TblKill As ListObject) As Integer
    'Works because each Ammo sheet's user input areas are actually Objects.
    'Uses Rge's row/column value relative to the sheet, subtracting TblKill's first row/column value
    Dim RgeProperty As Integer, TblStart As Integer
    If Data = "Enemy" Then
        RgeProperty = Rge.Row
        TblStart = TblKill.Range.Row
    ElseIf Data = "Level" Then
        RgeProperty = Rge.Column
        TblStart = TblKill.Range.Column
    Else
        MsgBox "Code error: wrong FindSelectionData argument. Process terminated."
        End
    End If
    
    FindSelectionData = RgeProperty - TblStart + 1  'Collections are indexed from 1.
End Function

Function FindLevelWeapons(Rge As Range, NewGamePlus As String, TblKill As ListObject)
    LevelWeapons = LevelName  'Distinguishes between the level from cell selection and the "level" being used for mWeaponsAvailableArray
    If NewGamePlus = "Yes" Then 
        
        If LevelWeapons = Level.OffshoreRig Then
            NewGamePlusORFlag = True  'Used later to alert user that they will start OR without weapons despite running NG+.
            Exit Function
        End If
        
        LevelWeapons = Level.NewGamePlus
        NewGamePlusORFlag = False
    End If
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

Function DeclareLevelArsenal(WeaponsAvailable As Variant)
    Dim i As Integer  'Incremented for loop
    Dim Size As Integer: Size = UBound(WeaponsAvailable, 2)  'Column count from WeaponsAvailable to size DeclareLevelArsenal array.
    
    ReDim DeclareLevelArsenal(1 To Size)
    For i = 1 To Size
        DeclareLevelArsenal(i) = WeaponsAvailable(LevelWeapons, i)
    Next i
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
=======
Attribute VB_Name = "mMainFunctions"
Option Explicit

Function FindItemIndex(Coll As Collection, Key As String) As Integer
    Dim i As Integer 'Loops over Coll to determine Key's repective index. VBA has no native procedure for this.
    For i = 1 To Coll.Count
        If Coll(i) = Coll(Key) Then
            Exit For  'Found it
        End If
    Next
    
    FindItemIndex = i
    
    If Coll(FindItemIndex) <> Coll(Key) Then  'Check if all items were processed but no match key exists for last value
        MsgBox "Coll(FindItemIndex) failed to = Coll(Key), process terminated."
        End
    End If
End Function

Function FindSelectionData(Data As String, Rge As Range, TblKill As ListObject) As Integer
    'Works because each Ammo sheet's user input areas are actually Objects.
    'Uses Rge's row/column value relative to the sheet, subtracting TblKill's first row/column value
    Dim RgeProperty As Integer, TblStart As Integer
    If Data = "Enemy" Then
        RgeProperty = Rge.Row
        TblStart = TblKill.Range.Row
    ElseIf Data = "Level" Then
        RgeProperty = Rge.Column
        TblStart = TblKill.Range.Column
    Else
        MsgBox "Code error: wrong FindSelectionData argument. Process terminated."
        End
    End If
    
    FindSelectionData = RgeProperty - TblStart + 1  'Collections are indexed from 1.
End Function

Function FindLevelWeapons(Rge As Range, NewGamePlus As String, TblKill As ListObject)
    LevelWeapons = LevelName  'Distinguishes between the level from cell selection and the "level" being used for mWeaponsAvailableArray
    If NewGamePlus = "Yes" Then
        
        If LevelWeapons = Level.OffshoreRig Then
            NewGamePlusORFlag = True  'Used later to alert user that they will start OR without weapons despite running NG+.
            Exit Function
        End If
        
        LevelWeapons = Level.NewGamePlus
        NewGamePlusORFlag = False
    End If
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

Function DeclareLevelArsenal(WeaponsAvailable As Variant)
    Dim i As Integer  'Incremented for loop
    Dim Size As Integer: Size = UBound(WeaponsAvailable, 2)  'Column count from WeaponsAvailable to size DeclareLevelArsenal array.
    
    ReDim DeclareLevelArsenal(1 To Size)
    For i = 1 To Size
        DeclareLevelArsenal(i) = WeaponsAvailable(LevelWeapons, i)
    Next i
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
>>>>>>> d7b9fa5... mWeaponsAvailable fixed, extra spaces removed
