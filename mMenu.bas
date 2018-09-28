Attribute VB_Name = "mMenu"
Sub ShowMeds()
    ThisWorkbook.Sheets("Meds").Visible = True
End Sub
Sub HideMeds()
    ThisWorkbook.Sheets("Meds").Visible = xlVeryHidden
End Sub
Sub ShowAmmoNone()
    ThisWorkbook.Sheets("Any%").Visible = xlVeryHidden
    ThisWorkbook.Sheets("Glitchless Any%").Visible = xlVeryHidden
    ThisWorkbook.Sheets("Secrets%").Visible = xlVeryHidden
    ThisWorkbook.Sheets("Glitchless Secrets%").Visible = xlVeryHidden
    ThisWorkbook.Sheets("100%").Visible = xlVeryHidden
    ThisWorkbook.Sheets("Glitchless 100%").Visible = xlVeryHidden
End Sub
Sub ShowAmmoAny()
    ShowAmmoNone
    ThisWorkbook.Sheets("Any%").Visible = True
End Sub
Sub ShowAmmoGlitchlessAny()
    ShowAmmoNone
    ThisWorkbook.Sheets("Glitchless Any%").Visible = True
End Sub
Sub ShowAmmoBothAny()
    ShowAmmoNone
    ThisWorkbook.Sheets("Any%").Visible = True
    ThisWorkbook.Sheets("Glitchless Any%").Visible = True
End Sub
'These respond to "Menu" sheet dropdowns when called from "ThisWorkbook" Module, change sheet visibilities as appropriate to user input

Sub ShowAmmoSecrets()
    ShowAmmoNone
    ThisWorkbook.Sheets("Secrets%").Visible = True
End Sub
Sub ShowAmmoGlitchlessSecrets()
    ShowAmmoNone
    ThisWorkbook.Sheets("Glitchless Secrets%").Visible = True
End Sub
Sub ShowAmmoBothSecrets()
    ShowAmmoNone
    ThisWorkbook.Sheets("Secrets%").Visible = True
    ThisWorkbook.Sheets("Glitchless Secrets%").Visible = True
End Sub
Sub ShowAmmo100()
    ShowAmmoNone
    ThisWorkbook.Sheets("100%").Visible = True
End Sub
Sub ShowAmmoGlitchless100()
    ShowAmmoNone
    ThisWorkbook.Sheets("Glitchless 100%").Visible = True
End Sub
Sub ShowAmmoBoth100()
    ShowAmmoNone
    ThisWorkbook.Sheets("100%").Visible = True
    ThisWorkbook.Sheets("Glitchless 100%").Visible = True
End Sub
Sub ShowAmmoAll()
    ThisWorkbook.Sheets("Any%").Visible = True
    ThisWorkbook.Sheets("Glitchless Any%").Visible = True
    ThisWorkbook.Sheets("Secrets%").Visible = True
    ThisWorkbook.Sheets("Glitchless Secrets%").Visible = True
    ThisWorkbook.Sheets("100%").Visible = True
    ThisWorkbook.Sheets("Glitchless 100%").Visible = True
End Sub
Sub ShowTableNone()
    ThisWorkbook.Sheets("100% Counts Glitched").Visible = xlVeryHidden
    ThisWorkbook.Sheets("100% Counts Glitchless").Visible = xlVeryHidden
End Sub
Sub ShowTableGlitched()
    ShowTableNone
    ThisWorkbook.Sheets("100% Counts Glitched").Visible = True
End Sub
Sub ShowTableGlitchless()
    ShowTableNone
    ThisWorkbook.Sheets("100% Counts Glitchless").Visible = True
End Sub
Sub ShowTableBoth()
    ThisWorkbook.Sheets("100% Counts Glitched").Visible = True
    ThisWorkbook.Sheets("100% Counts Glitchless").Visible = True
End Sub
