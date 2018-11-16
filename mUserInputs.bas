Attribute VB_Name = "mUserInputs"
Option Explicit

Function KillInput(TotalKills As Integer, Weapon As Collection, _
                   EnemySelect As Integer, LvlSelect As Integer, WeaponIndex As Integer, _
                   Rge As Range, tblShots As ListObject) As Integer

    Dim Check As Boolean
    Dim Kills As Integer

    Do  'Loop to prompt user to indicate weapon choice for each of the enemies killed
        
        Check = True
        Kills = InputBox("Enter Number of " & Weapon("Weapon" & WeaponIndex) & " Kills", Weapon(WeaponIndex) & "Input", 0)
        
        If StrPtr(Kills) = 0 Then  'If user cancels:
            
            If MsgBox("Cancelling the weapon prompt sets the enemy selected's kill count to 0 for this level, and resets any ammo used for the enemy to 0. If this is what you want, hit Cancel, otherwise, hit Retry.", vbQuestion + vbRetryCancel, "Confirm Cancellation") = vbCancel Then
                Debug.Print "User cancelled from Kills InputBox."
                ActiveCell.Value = 0
                '!!! Set all relevant weapon-kill archive values to 0.
                End
            End If
        
        ElseIf IsNumeric(Kills) = False Then  'If user inputs non-number:
            MsgBox "Input must be a whole number.", vbCritical + vbOKOnly, "Warning"
            Check = False
        
        ElseIf Val(Kills) <> Val(Round(Kills)) Then  'If user inputs a non-integer:
            MsgBox "Whole numbers only!", vbCritical + vbOKOnly, "Warning"
            Check = False
       
        ElseIf Val(Kills) < 0 Then  'If user inputs a number either negative or above total indicated:
            MsgBox "Input must be non-negative.", vbCritical + vbOKOnly, "Negative Number"
            Check = False
            
        ElseIf Val(Kills) > Rge.Value Then
            MsgBox "Number must be less than value indicated in cell dropdown.", vbCritical + vbOKOnly, "Value Discrepancy"
            Check = False
            
        ElseIf Val(Kills) + TotalKills > Rge.Value Then
            MsgBox TotalKills & " kills have already been assigned, only " & Rge.Value - TotalKills & " kills remain.", vbCritical + vbOKOnly, "Total Kills Discrepancy"
            Check = False
            
        End If

    Loop Until Check = True
    
    'Upon successful user input, update archive cell value
    Dim Shots As Integer
    Shots = ShotsToKill(EnemySelect, WeaponIndex, tblShots)
    
    Dim Name As String: Name = Weapon("Weapon" & WeaponIndex) & " Kills"
    Dim Column As Long: Column = CollapseRun()
    Dim Row As Long: Row = CollapseEnemyAndLevel(Rge)
    
    Dim Cell As Range: Set Cell = ThisWorkbook.Worksheets(Name).Cells(Row, Column)
    Dim Value As Integer: Value = Kills * Shots
    
    ThisWorkbook.Worksheets(Name).Unprotect
    Cell.Value = Value
    ThisWorkbook.Worksheets(Name).Protect
    
    KillInput = Kills
End Function

Function ShotsToKill(EnemySelect As Integer, WeaponIndex As Integer, tblShots As ListObject)
    
    'Getting column data to bound loop
    Dim ColStart As Integer: ColStart = tblShots.Range.Column
    Dim ColEnd As Integer: ColEnd = ColStart + (tblShots.ListColumns.Count - 1)
    
    'Base the row off of EnemySelect
    Dim Row As Long
    Row = tblShots.Range.Row + (EnemySelect - 1)
    
    'Can use color to compare since sheet color is pre-formatted, assuming user won't change these
    Dim Color As Long: Color = Cells(Row, ColStart).Interior.Color
    Dim Count As Integer: Count = 1
    Dim Column As Integer
    
    Dim y As Integer
    For y = ColStart To ColEnd
        
        If Cells(Row, y).Interior.Color <> Color Then
            
            If Count = WeaponIndex Then
                Column = y - 1
                Exit For
            End If
            Count = Count + 1
            
        End If
        
    Next y
    
    ShotsToKill = Cells(Row, Column).Value
End Function
