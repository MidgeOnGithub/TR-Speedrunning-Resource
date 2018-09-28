Attribute VB_Name = "mTables"
'Subroutine to determine whether or not to add inaccessible WotMD shark

Sub SharkKillToggle()

'Declare variables
Dim SharkKill As Boolean, SharkRng As Range

'C11 is the WotMD Kills cell
Set SharkRng = ThisWorkbook.ActiveSheet.Range("C11")

SharkKill = Application.InputBox(Prompt:="Are you going to kill the optional shark? Leaderboard rules do not require it, so unless you want to do extra work, leave the default option. Switch to something Excel interprets as 'True' if you wish to kill the shark.", _
            Title:="Shark Kill Prompt", Default:="False", Type:="4")

'Change WotMD kills value based on user input
If SharkKill = True Then
    ThisWorkbook.ActiveSheet.Unprotect
    SharkRng.Value = 36
    ThisWorkbook.ActiveSheet.Protect
Else
    ThisWorkbook.ActiveSheet.Unprotect
    SharkRng.Value = 35
    ThisWorkbook.ActiveSheet.Protect
End If

End Sub


