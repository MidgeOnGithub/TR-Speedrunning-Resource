VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "s04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim strRun As String
    strRun = FindRunType(Target.Worksheet)
    
    Dim tblKills As ListObject
    Set tblKills = ActiveSheet.ListObjects("tbl" + strRun + "Kills")
    
    'Secrets row is directly after the end of tblKills, thus adding row count works
    Dim SecretsRow As Long
    SecretsRow = tblKills.Range.Row + tblKills.Range.Rows.Count
    
    If Not Intersect(Target, tblKills.DataBodyRange) Is Nothing Then
        
        Application.EnableEvents = False
        Call Main  'Call Main if user changed one of the kill dropdowns
        Application.EnableEvents = True
        
    ElseIf Target.Row = SecretsRow Then
        
        Dim tblCol As Long
        tblCol = tblKills.Range.Column
        
        'This If is Specific to TR2, checking if they change secrets on levels where doing so affects weapon availability
        If Target.Column = tblCol Or Target.Column = tblCol + 4 Or Target.Column = tblCol + 7 Then
            
            Dim strLevels As String, strWeapons As String
            
            Select Case Target.Column
            
                Case tblCol
                    strLevels = "Great Wall through Opera House"
                    strWeapons = "Grenade Launcher"
                    
                Case tblCol + 4
                    strLevels = "Offshore Rig"
                    strWeapons = "Uzi"
                    
                Case tblCol + 7
                    strLevels = "Maria Doria and Living Quarters"
                    strWeapons = "Grenade Launcher"
            
            End Select
            
            MsgBox strWeapons & " availability changes with this secret." & vbCr & vbCr & "If you changed this cell from Yes to No or vice versa," & vbCr & "and you previously used affected weapons for kills," & vbCr & "re-input enemy kills for " & strLevels & vbCr & "to be re-prompted for weapon-kill assignments.", vbInformation + vbOKOnly, "Weapon Availability Change Notice"
            
        End If
        
    End If
    
End Sub
