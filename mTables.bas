Attribute VB_Name = "mTables"
Option Explicit

<<<<<<< HEAD
Sub SharkKillToggle() 'Sub determines whether or not to add inaccessible WotMD shark.
    Dim SharkRge As Range
    Set SharkRge = ThisWorkbook.ActiveSheet.Range("C11") 'C11 is the WotMD Kills cell.
    
    'Change WotMD kills value based on user input.
    If MsgBox("Are you going to kill the optional shark?" & vbCr & "Leaderboard rules do not require it.", _
              vbQuestion + vbYesNo, "Shark Kill Prompt") = vbYes Then
        ThisWorkbook.ActiveSheet.Unprotect
        SharkRge.Value = 36
        ThisWorkbook.ActiveSheet.Protect
        
    Else 'User selected vbNo.
        ThisWorkbook.ActiveSheet.Unprotect
        SharkRge.Value = 35
        ThisWorkbook.ActiveSheet.Protect
        
    End If

End Sub

Sub MedPickupToggle()
    Dim MedRge As Range
    Set MedRge = ThisWorkbook.ActiveSheet.Range("D18") 'D18 is the Temple of Xian Pickups cell.
    
    'Change Xian pickups value based on user input.
    If MsgBox("Are you picking up the previously unobtainable large med?", _
              vbQuestion + vbYesNo, "Med Pickup Prompt") = vbYes Then
        ThisWorkbook.ActiveSheet.Unprotect
        MedRge.Value = 40
        ThisWorkbook.ActiveSheet.Protect
        
    Else 'User selected vbNo.
        ThisWorkbook.ActiveSheet.Unprotect
        MedRge.Value = 39
        ThisWorkbook.ActiveSheet.Protect
        
    End If
End Sub

=======
'This issue is specific to Tomb Raider II
Sub SharkKillToggle() 'Determine whether or not to add inaccessible WotMD shark
    Dim SharkRng As Range: Set SharkRng = ThisWorkbook.ActiveSheet.Range("C11")  'C11 is the WotMD Kills cell
    
    ThisWorkbook.ActiveSheet.Unprotect
    If MsgBox("Are you going to kill the optional shark? Leaderboard rules do not require it.", vbQuestion + vbYesNo, "Shark Kill Prompt") = vbYes Then
        SharkRng.Value = 36
    Else  'Changes WotMD Kills value based on user input
        SharkRng.Value = 35
    End If 
    ThisWorkbook.ActiveSheet.Protect
End Sub
>>>>>>> f8de6bc... Minor progress with mUserInputs and mPublicFunctions, some code beautification
