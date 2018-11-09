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
