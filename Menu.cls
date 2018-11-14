VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    'For the top menu dropdown, "Info"
    Select Case Range("Info")
        Case "Yes"
            Call ShowSheet(True, "Information")
        
        Case "No"
            Call ShowSheet(False, "Information")
    End Select
    
    'For the top menu dropdown, "Meds."
    Select Case Range("Meds")
        Case "Yes"
            Call ShowSheet(True, "Meds")
        
        Case "No"
            Call ShowSheet(False, "Meds")
    End Select
    
    'For the next menu dropdown, "Ammo"
    Select Case Range("Ammo")
        Case "None"
            Call HideAmmo
        
        Case "Any%"
            Call HideAmmo
            Call ShowSheet(True, "Any%")
        
        Case "Glitchless Any%"
            Call HideAmmo
            Call ShowSheet(True, "Glitchless Any%")
        
        Case "Both Any%"
            Call HideAmmo
            Call ShowSheet(True, "Any%")
            Call ShowSheet(True, "Glitchless Any%")
        
        Case "Secrets%"
            Call HideAmmo
            Call ShowSheet(True, "Secrets%")
        
        Case "Glitchless Secrets%"
            Call HideAmmo
            Call ShowSheet(True, "Glitchless Secrets%")
        
        Case "Both Secrets%"
            Call HideAmmo
            Call ShowSheet(True, "Secrets%")
            Call ShowSheet(True, "Glitchless Secrets%")
        
        Case "1"
            Call HideAmmo
            Call ShowSheet(True, "100%")
        
        Case "Glitchless 100%"
            Call HideAmmo
            Call ShowSheet(True, "Glitchless 100%")
        
        Case "Both 100%"
            Call HideAmmo
            Call ShowSheet(True, "100%")
            Call ShowSheet(True, "Glitchless 100%")
        
        Case "All"
            Call ShowSheet(True, "Any%")
            Call ShowSheet(True, "Glitchless Any%")
            Call ShowSheet(True, "Secrets%")
            Call ShowSheet(True, "Glitchless Secrets%")
            Call ShowSheet(True, "100%")
            Call ShowSheet(True, "Glitchless 100%")
    End Select
    
    'For the next menu dropdown, "Flares"
    'Select Case Range("Flares")
        'Case "Yes"
            'Call ShowSheet(True, "Flares")
            
        'Case "No"
            'Call ShowSheet(False, "Flares")
    'End Select
    
    'For the last menu dropdown, "Table"
    Select Case Range("Table")
        Case "None"
            HideTable
            
        Case "Glitched"
            HideTable
            Call ShowSheet(True, "100% Counts")
            
        Case "Glitchless"
            HideTable
            Call ShowSheet(True, "100% Counts Glitchless")
            
        Case "Both"
            HideTable
            Call ShowSheet(True, "100% Counts")
            Call ShowSheet(True, "100% Counts Glitchless")
    End Select

End Sub

Private Function ShowSheet(boolCase As Boolean, SheetsToShow As String) 'Sheets are named the same as possible input strings.

    If boolCase = True Then
        ThisWorkbook.Sheets(SheetsToShow).Visible = True
    Else
        ThisWorkbook.Sheets(SheetsToShow).Visible = xlVeryHidden
    End If
    
End Function

Private Function HideAmmo()

With ThisWorkbook

    .Sheets("Any%").Visible = xlVeryHidden
    .Sheets("Glitchless Any%").Visible = xlVeryHidden
    .Sheets("Secrets%").Visible = xlVeryHidden
    .Sheets("Glitchless Secrets%").Visible = xlVeryHidden
    .Sheets("100%").Visible = xlVeryHidden
    .Sheets("Glitchless 100%").Visible = xlVeryHidden

End With

End Function

Private Function HideTable()
    
    ThisWorkbook.Sheets("100% Counts").Visible = xlVeryHidden
    ThisWorkbook.Sheets("100% Counts Glitchless").Visible = xlVeryHidden
    
End Function
