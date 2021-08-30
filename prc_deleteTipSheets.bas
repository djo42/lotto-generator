Attribute VB_Name = "prc_deleteTipSheets"
Sub deleteTipSheets()

    Dim ws As Worksheet
 
    For Each ws In ThisWorkbook.Worksheets
    
        If ws.Name <> "GameData" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
        
    Next
    
End Sub

