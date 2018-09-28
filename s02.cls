VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "s02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

'For the top menu dropdown, cell D2/"Meds"
Select Case Range("Meds")
    Case "Yes": ShowMeds
    Case "No": HideMeds
End Select

'For the next menu dropdown, cell D3/"Ammo"
Select Case Range("Ammo")
    Case "None": ShowAmmoNone
    Case "Any%": ShowAmmoAny
    Case "Glitchless Any%": ShowAmmoGlitchlessAny
    Case "Both Any%": ShowAmmoBothAny
    Case "Secrets%": ShowAmmoSecrets
    Case "Glitchless Secrets%": ShowAmmoGlitchlessSecrets
    Case "Both Secrets%": ShowAmmoBothSecrets
    Case "1": ShowAmmo100
    Case "Glitchless 100%": ShowAmmoGlitchless100
    Case "Both 100%": ShowAmmoBoth100
    Case "All": ShowAmmoAll
End Select

'For the next menu dropdown, cell D4/"Flares"
'Select Case Range("Flares")
    'Case "Yes": ShowFlares
    'Case "No": HideFlares
    
'For the last menu dropdown, cell D5/"Table"
Select Case Range("Table")
    Case "None": ShowTableNone
    Case "Glitched": ShowTableGlitched
    Case "Glitchless": ShowTableGlitchless
    Case "Both": ShowTableBoth
End Select

End Sub
