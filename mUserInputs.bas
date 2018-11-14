Attribute VB_Name = "mUserInputs"
Option Explicit

Function KillInput(Level As Collection, Weapon As Collection, _
                   EnemySelect As Integer, LvlSelect As Integer, _
                   Rge As Range, Sht As Worksheet, _
                   RunType As String, strRulesCategory As String, _
                   TblKills As ListObject, WeaponIndex As Variant)

    Dim Check As Boolean
    
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
            Debug.Print "Non-number."
        
        ElseIf Val(Kills) <> Val(Round(Kills)) Then  'If user inputs a non-integer:
            MsgBox "Whole numbers only!", vbCritical + vbOKOnly, "Warning"
            Check = False
            Debug.Print "Non (literal) integer."
       
        ElseIf (Val(Kills) < 0 Or Val(Kills) > Rge.Value) Then  'If user inputs a number either negative or above total indicated:
            MsgBox "Input Must Be Non-Negative and <= value entered into cell.", vbCritical + vbOKOnly, "Number Not Logical"
            Check = False
            Debug.Print "Input was negative or greater than cell-entered Kills."
        End If

    Loop Until Check = True

    'Need to use functions to get archive sheet cell to input the kill!!!
    Dim ArchiveColumn As Long
    Set ArchiveColumn = CollapseRun()

    Dim ArchiveRow As Long
    Set ArchiveRow = CollapseEnemyAndLevel()

    'Upon successful user input, update archive sheet value
    ThisWorkbook.Worksheets(strRulesCategory & " " & RunType & "%" & Weapon("Weapon" & WeaponIndex) & " Kills").Range(ArchiveRow, ArchiveColumn).Value = Kills
End Function

Function ShotsToKill(EnemySelect As Integer, LvlSelect As Integer, TblKills As ListObject)
 'Is this needed???
End Function
