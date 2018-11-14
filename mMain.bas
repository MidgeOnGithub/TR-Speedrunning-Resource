<<<<<<< HEAD
<<<<<<< HEAD
Attribute VB_Name = "mMain"
Option Explicit

Sub Main()
    'Shorthands for easier coding. ======================================
    Dim Rge As Range: Set Rge = ActiveCell
    Dim Sht As Worksheet: Set Sht = ThisWorkbook.ActiveSheet

    'Needed to correctly name Tbl variables. --------
    Dim RulesCategory As Boolean: RulesCategory = IsGlitchless(Sht)
    Dim strRulesCategory As String, RunType As String

    If RulesCategory = True Then
        strRulesCategory = "Glitchless"
    Else
        strRulesCategory = "" 'Glitched; leaving blank for sheet name inputs later.
    End If
    RunType = FindRunType(Sht)
    '-----------------------------------------------|

    Dim TblShots As ListObject: Set TblShots = Sht.ListObjects("tbl" & RunType & "Shots")
    Dim TblKill As ListObject: Set TblKill = Sht.ListObjects("tbl" & RunType & "Kills")
    '===================================================================|

    'Uses mPopulate module. =============================================
    Dim Enemy As Collection: Set Enemy = PopulateColl("Enemy")
    Dim Level As Collection: Set Level = PopulateColl("Level")
    Dim Weapon As Collection: Set Weapon = PopulateColl("Weapon")
    '==================================================================||

    'Finds collection index values for selected cell. ===================
    Dim EnemySelect As Integer: EnemySelect = FindSelectionData("Enemy", Rge, TblKill)
    Dim LvlSelect As Integer: LvlSelect = FindSelectionData("Level", Rge, TblKill)
    '=================================================================|||
    
    'Finds other run info. ==============================================
    Dim NewGamePlus As Boolean: NewGamePlus = IsNewGamePlus(Level, LvlSelect, Sht)
    'Uses mWeaponsAvailable module.
    Dim LevelArsenal() As Variant: LevelArsenal = ArrayLevelWeapons(Level, LvlSelect, NewGamePlus, RunType, Sht, Weapon)
    '================================================================||||
    
    'No prompts needed if user changes enemy kill count to 0, just adjust cell(s)/formula(s) for each weapon affected.
    If Val(Rge.Value) = 0 Then
        ThisWorkbook.Worksheets(RunType & "% " & "Pistol Kill Counts").Range("B5").Value = 0
        '!!! Need to edit to appropriately adjust all applicable weapons' ammo formulas to 0
        Exit Sub
    End If
    
    'Uses mUserInputs module. ============================================
    'Start a loop beginning with the first weapon available in the level; step through rest of available weapons
    'Do
        'PistolsKillInput
        'If Val(Pistols) = Val(Cell.Value) Then
            'Debug.Print "Pistols kills equalled total kill count for EnemyName, no more prompts needed."
            'GoTo Formulas
        'End If
        'ShotgunKillInput
        'If Val(Pistols) + Val(Shotgun) = Val(Cell.Value) Then
            'Debug.Print "Pistols + shotgun kills equalled total kill count for EnemyName, no more prompts needed."
            'GoTo Formulas
        'End If
        'Other weapon subroutines will go here
    'Loop Until Val(Pistols) + Val(Autos) + Val(Shotgun) + Val(Uzis) + Val(HarpoonGun) + Val(M16) + Val(GrenadeLauncher) = Val(Cell.Value)
    '===============================================================|||||
    
    'Formulas ???
End Sub
=======
Attribute VB_Name = "mMain"
Option Explicit

Sub Main()
    'Shorthands for easier coding. ======================================
    Dim Rge As Range: Set Rge = ActiveCell
    Dim Sht As Worksheet: Set Sht = ThisWorkbook.ActiveSheet

    'Needed to correctly name Tbl variables. --------
    Dim RulesCategory As Boolean: RulesCategory = IsGlitchless(Sht)
    Dim strRulesCategory As String, RunType As String

    If RulesCategory = True Then
        strRulesCategory = "Glitchless"
    Else
        strRulesCategory = "" 'Glitched; leaving blank for sheet name inputs later.
    End If
    RunType = FindRunType(Sht)
    '-----------------------------------------------|

    Dim TblShots As ListObject: Set TblShots = Sht.ListObjects("tbl" & RunType & "Shots")
    Dim TblKill As ListObject: Set TblKill = Sht.ListObjects("tbl" & RunType & "Kills")
    '===================================================================|

    'Uses mPopulate module. =============================================
    Dim Enemy As Collection: Set Enemy = PopulateColl("Enemy")
    Dim Level As Collection: Set Level = PopulateColl("Level")
    Dim Weapon As Collection: Set Weapon = PopulateColl("Weapon")
    '==================================================================||

    'Finds collection index values for selected cell. ===================
    Dim EnemySelect As Integer: EnemySelect = FindSelectionData("Enemy", Rge, TblKill)
    Dim LvlSelect As Integer: LvlSelect = FindSelectionData("Level", Rge, TblKill)
    '=================================================================|||
    
    'Finds other run info. ==============================================
    Dim NewGamePlus As Boolean: NewGamePlus = IsNewGamePlus(Level, LvlSelect, Sht)
    'Uses mWeaponsAvailable module.
    Dim LevelArsenal() As Variant: LevelArsenal = LevelWeapons(Level, LvlSelect, NewGamePlus, RunType, Sht, Weapon)
    '================================================================||||
    
    'No prompts needed if user changes enemy kill count to 0, just adjust cell(s)/formula(s) for each weapon affected.
    If Val(Rge.Value) = 0 Then
        ThisWorkbook.Worksheets(RunType & "% " & "Pistol Kill Counts").Range("B5").Value = 0
        '!!! Need to edit to appropriately adjust all applicable weapons' ammo formulas to 0
        Exit Sub
    End If
    
    'Uses mUserInputs module. ===========================================
    Dim i As Integer: i = 1
    Dim TotalKills As Integer: Kill = 0
    
    Do
        
        If LevelArsenal(i) = 1 Then
             TotalKills = TotalKills + KillInput(TotalKills, Weapon, EnemySelect, LvlSelect, Rge, Sht, TblKills, WeaponIndex)
        End If
        
        If TotalKills = Rge.Value Then Exit Do
        
        i = i + 1
        
    Loop Until (i > Weapon.Count)
    '===============================================================|||||
    
End Sub
>>>>>>> efdc547... mMain workflow finished
=======
Attribute VB_Name = "mMain"
Option Explicit

Sub Main()
    'Shorthands for easier coding. ======================================
    Dim Rge As Range: Set Rge = ActiveCell
    Dim Sht As Worksheet: Set Sht = ThisWorkbook.ActiveSheet

    'Needed to correctly name Tbl variables. --------
    Dim RulesCategory As Boolean: RulesCategory = IsGlitchless(Sht)
    Dim strRulesCategory As String, RunType As String

    If RulesCategory = True Then
        strRulesCategory = "Glitchless"
    Else
        strRulesCategory = "" 'Glitched; leaving blank for sheet name inputs later.
    End If
    RunType = FindRunType(Sht)
    '-----------------------------------------------|
    
    Dim tblAmmo As ListObject: Set tblAmmo = Sht.ListObjects("tbl" & RunType & "Ammo")
    Dim tblKills As ListObject: Set tblKills = Sht.ListObjects("tbl" & RunType & "Kills")
    Dim tblShots As ListObject: Set tblShots = Sht.ListObjects("tbl" & RunType & "Shots")
    '===================================================================|

    'Uses mPopulate module. =============================================
    Dim Enemy As Collection: Set Enemy = PopulateColl("Enemy")
    Dim Level As Collection: Set Level = PopulateColl("Level")
    Dim Weapon As Collection: Set Weapon = PopulateColl("Weapon")
    '==================================================================||

    'Finds collection index values for selected cell. ===================
    Dim EnemySelect As Integer: EnemySelect = FindSelectionData("Enemy", Rge, tblKills)
    Dim LvlSelect As Integer: LvlSelect = FindSelectionData("Level", Rge, tblKills)
    '=================================================================|||
    
    'Finds other run info. ==============================================
    Dim NewGamePlus As Boolean: NewGamePlus = IsNewGamePlus(Level, LvlSelect, Sht)
    'Uses mWeaponsAvailable module.
    Dim LevelArsenal() As Variant: LevelArsenal = LevelWeapons(Level, LvlSelect, NewGamePlus, RunType, Sht, Weapon)
    '================================================================||||
    
    'No prompts needed if user changes enemy kill count to 0, just adjust cell(s)/formula(s) for each weapon affected.
    If Val(Rge.Value) = 0 Then
        '!!! Need to edit to appropriately adjust all applicable weapons' ammo formulas to 0
        Exit Sub
    End If
    
    'Uses mUserInputs module. ===========================================
    Dim WeaponIndex As Integer: WeaponIndex = 1
    Dim TotalKills As Integer: TotalKills = 0
    
    Do Until WeaponIndex = Weapon.Count
        
        If LevelArsenal(WeaponIndex) = 1 Then
             TotalKills = TotalKills + KillInput(TotalKills, Weapon, EnemySelect, LvlSelect, WeaponIndex, Rge, tblShots)
        End If

        If TotalKills = Rge.Value Then Exit Do
        
        WeaponIndex = WeaponIndex + 1
        
    Loop
    '===============================================================|||||
    
    'Force the recalculation of ammo values once Main's work has finished
    'Needed because of the use of NVIndirect, a static function
    tblAmmo.ListColumns(LvlSelect).Range.Calculate
    
End Sub
>>>>>>> 111d2d8... Finished Main, s04 + custom function issues
