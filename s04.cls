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
    strRun = FindRunType(ActiveSheet)
    
    Dim tblKills As ListObject
    Set tblKills = ActiveSheet.ListObjects("tbl" + strRun + "Kills")
    
    Dim Row As Long
    Row = tblKills.Range.Row + tblKills.Range.Rows.Count
    
    If Not Intersect(Target, tblKills.DataBodyRange) Is Nothing Then
        
        Call Main  'Call Main if user changed one of the kill dropdowns
    
    ElseIf Target.Row = Row Then
        
        Dim tblAmmo As ListObject
        Set tblAmmo = ActiveSheet.ListObjects("tbl" + strRun + "Ammo")
        
        'If the user changes a secrets dropdown, force the ammo values to recalculate
        Dim Column As Long
        Column = Target.Column - tblAmmo.Range.Column + 1
        
        tblAmmo.ListColumns(Column).Range.Calculate
        
    End If
    
End Sub
