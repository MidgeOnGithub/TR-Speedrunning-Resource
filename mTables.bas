Attribute VB_Name = "mTables"
Option Explicit

Sub SharkKillToggle() 'Determine whether or not to add inaccessible WotMD shark
    Dim SharkRng As Range: Set SharkRng = ThisWorkbook.ActiveSheet.Range("C11") 'C11 is the WotMD Kills cell
    
    ThisWorkbook.ActiveSheet.Unprotect
    If MsgBox("Are you going to kill the optional shark? Leaderboard rules do not require it, so unless you want to do extra work, leave the default option. Switch to anything other than 'Yes' if you wish to kill the shark.", vbQuestion + vbYesNo, "Shark Kill Prompt") = vbYes Then
        SharkRng.Value = 36
    Else 'Changes WotMD kills value based on user input
        SharkRng.Value = 35
    End If 
    ThisWorkbook.ActiveSheet.Protect
End Sub