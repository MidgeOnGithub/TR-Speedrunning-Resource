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

'For the top menu dropdown, cell D2/Meds.
Select Case Range("Meds")
    
    Case "Yes": ShowMeds (True)
    
    Case "No": ShowMeds (False)
    
End Select

'For the next menu dropdown, cell D3/Ammo.
Select Case Range("Ammo")
    
    Case "None": ShowAmmoNone
    
    Case "Any%": ShowAmmoNone
        ShowAmmo ("Any%")
    
    Case "Glitchless Any%": ShowAmmoNone
        ShowAmmo ("Glitchless Any%")
    
    Case "Both Any%": ShowAmmoNone
        ShowAmmo ("Any%")
        ShowAmmo ("Glitchless Any%")
    
    Case "Secrets%": ShowAmmoNone
        ShowAmmo ("Secrets%")
    
    Case "Glitchless Secrets%": ShowAmmoNone
        ShowAmmo ("Glitchless Secrets%")
    
    Case "Both Secrets%": ShowAmmoNone
        ShowAmmo ("Secrets%")
        ShowAmmo ("Glitchless Secrets%")
    
    Case "1": ShowAmmoNone
        ShowAmmo ("100%")
    
    Case "Glitchless 100%": ShowAmmoNone
        ShowAmmoNone
        ShowAmmo ("Glitchless 100%")
    
    Case "Both 100%": ShowAmmoNone
        ShowAmmo ("100%")
        ShowAmmo ("Glitchless 100%")
    
    Case "All"
        ShowAmmo ("Any%")
        ShowAmmo ("Glitchless Any%")
        ShowAmmo ("Secrets%")
        ShowAmmo ("Glitchless Secrets%")
        ShowAmmo ("100%")
        ShowAmmo ("Glitchless 100%")

End Select

'For the next menu dropdown, cell D4/Flares.
'Select Case Range("Flares")
    'Case "Yes": ShowFlares
    'Case "No": HideFlares
    
'For the last menu dropdown, cell D5/Table.
Select Case Range("Table")
    
    Case "None"
        ShowTableNone
        
    Case "Glitched"
        ShowTableNone
        ShowTable ("100% Counts")
        
    Case "Glitchless"
        ShowTableNone
        ShowTable ("100% Counts Glitchless")
        
    Case "Both"
        ShowTableNone
        ShowTable ("100% Counts")
        ShowTable ("100% Counts Glitchless")

End Select

End Sub

Private Function ShowMeds(boolCase As Boolean)

If boolCase = True Then
    ThisWorkbook.Sheets("Meds").Visible = True
    
Else
    ThisWorkbook.Sheets("Meds").Visible = xlVeryHidden
    
End If
    
End Function

Private Function ShowAmmoNone()

With ThisWorkbook

    .Sheets("Any%").Visible = xlVeryHidden
    .Sheets("Glitchless Any%").Visible = xlVeryHidden
    .Sheets("Secrets%").Visible = xlVeryHidden
    .Sheets("Glitchless Secrets%").Visible = xlVeryHidden
    .Sheets("100%").Visible = xlVeryHidden
    .Sheets("Glitchless 100%").Visible = xlVeryHidden

End With

End Function

Private Function ShowAmmo(SheetsToShow As String)

ThisWorkbook.Sheets(SheetsToShow).Visible = True 'Sheets are named the same as possible input strings.

End Function

Private Function ShowTableNone()
    
ThisWorkbook.Sheets("100% Counts").Visible = xlVeryHidden
ThisWorkbook.Sheets("100% Counts Glitchless").Visible = xlVeryHidden
    
End Function

Private Function ShowTable(SheetsToShow As String)

ThisWorkbook.Sheets(SheetsToShow).Visible = True 'Sheets are named the same as possible input strings.

End Function
