Attribute VB_Name = "mWeaponsAvailable"
Option Explicit

'This function returns to its caller a pseudo-Boolean array telling what weapons are available in a level
Function LevelWeapons(Level As Collection, LvlSelect As Integer, _
                           NewGamePlus As Boolean, RunType As String, Sht As Worksheet, _
                           Weapon As Collection) As Variant

    'Hardcoded for TR2.
    'Reading this 2D array/matrix, it's 18 rows with 7 columns; 17 of the rows represent levels GW - DL and the last represents levels where all weapons are given when running NG+.
    'Home Sweet Home is disregarded as weapon choice is predetermined.
    '0 means the weapon is unavailable, 1 means available.
    'Some weapons are only available in levels if you collect all secrets; Ifs at the bottom handle these.
    Dim ArrWeaponsAvailable As Variant
    ArrWeaponsAvailable = [{1,1,0,0,0,0,0;1,1,1,0,0,0,0;1,1,1,1,0,0,0;1,1,1,1,0,0,0;1,1,1,0,1,0,0;1,1,1,1,1,1,0;1,1,1,1,1,1,0;1,1,1,1,1,1,0;1,1,1,1,1,1,0;1,1,1,1,1,1,1;1,1,1,1,1,1,1;1,1,1,1,1,1,1;1,1,1,1,1,1,1;1,1,1,1,1,1,1;1,1,1,1,1,1,1;1,1,1,1,1,1,1;1,1,1,1,1,1,1;1,1,1,1,1,1,1}]
    If Not ((Level(LvlSelect) = Level("Level1")) Or (Level(LvlSelect) = Level("Level5")) Or (Level(LvlSelect) = Level("Level8"))) Then
        GoTo WhichRow 'If the runner is not running a level where having all secrets affects weapon availability, skip ahead
    End If

    'Get ranges of cells for levels where All Secrets toggle affects weapon availability in NG Any% runs.
    Dim GWSecrets As Range: Set GWSecrets = Sht.Range("GWSecretsCheck")
    Dim ORSecrets As Range: Set ORSecrets = Sht.Range("ORSecretsCheck")
    Dim MDSecrets As Range: Set MDSecrets = Sht.Range("MDSecretsCheck")
    
    'If the player is running Secrets% or 100%, switch all values.
    If Not RunType = "Any" Then
        Call GWWeaponsAvailable(ArrWeaponsAvailable, Level, Weapon)
        Call ORWeaponsAvailable(ArrWeaponsAvailable, Level, Weapon)
        Call MDWeaponsAvailable(ArrWeaponsAvailable, Level, Weapon)
        Exit Function
    End If
    
    'Check if Any% runner is collecting secrets in levels of concern, adjust values as necessary.
    If GWSecrets.Value = "Yes" Then
        Call GWWeaponsAvailable(ArrWeaponsAvailable, Level, Weapon)
    End If
    If ORSecrets.Value = "Yes" Then
        Call ORWeaponsAvailable(ArrWeaponsAvailable, Level, Weapon)
    End If
    If MDSecrets.Value = "Yes" Then
        Call MDWeaponsAvailable(ArrWeaponsAvailable, Level, Weapon)
    End If

WhichRow:
    Dim Row As Integer
    If NewGamePlus = True Then
        Row = 18
    Else
        Row = LvlSelect
    End If
    
    'Cannot directly set a 1D array equal to one row of a 2D array, apparently, so need to have "holding array" before giving actual return array
    ReDim ReturnArray(1 To Weapon.Count)
    Dim i As Integer
    For i = 1 To Weapon.Count
        ReturnArray(i) = ArrWeaponsAvailable(Row, i)
    Next i
    
    LevelWeapons = ReturnArray
    Set ArrWeaponsAvailable = Nothing 'Redundant with VBA End Function; clears memory.
End Function

'Following functions set specific weapons in specific levels to be available if user has indicated they are collecting all secrets in certain levels.
Private Function GWWeaponsAvailable(ArrWeaponsAvailable As Variant, Level As Collection, Weapon As Collection)
    ArrWeaponsAvailable(FindItemIndex(Level, "Level1"), FindItemIndex(Weapon, "Weapon7")) = 1
    ArrWeaponsAvailable(FindItemIndex(Level, "Level2"), FindItemIndex(Weapon, "Weapon7")) = 1
    ArrWeaponsAvailable(FindItemIndex(Level, "Level3"), FindItemIndex(Weapon, "Weapon7")) = 1
    ArrWeaponsAvailable(FindItemIndex(Level, "Level4"), FindItemIndex(Weapon, "Weapon7")) = 1
End Function

Private Function ORWeaponsAvailable(ArrWeaponsAvailable As Variant, Level As Collection, Weapon As Collection)
    ArrWeaponsAvailable(FindItemIndex(Level, "Level5"), FindItemIndex(Weapon, "Weapon4")) = 1
End Function

Private Function MDWeaponsAvailable(ArrWeaponsAvailable As Variant, Level As Collection, Weapon As Collection)
    ArrWeaponsAvailable(FindItemIndex(Level, "Level8"), FindItemIndex(Weapon, "Weapon7")) = 1
    ArrWeaponsAvailable(FindItemIndex(Level, "Level9"), FindItemIndex(Weapon, "Weapon7")) = 1
End Function

Function FindItemIndex(Coll As Collection, Key As String) As Integer
    Dim i As Integer  'Loops over Coll to determine Key's repective index. VBA has no native procedure for this.
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
