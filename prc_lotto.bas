Attribute VB_Name = "prc_lotto"
Option Explicit

Sub lotto()

    Dim sheetName, namePart As String
    Dim wb As Workbook
    Dim ws, gd, item As Variant
    Dim c, i, m, n, r, t As Long
    Dim e As Range
    Dim y As New Collection
    Dim z As New Collection
     
    Set wb = ThisWorkbook
    Set gd = Sheets(1)
    
    sheetName = GUID(True, False)

    Set ws = wb.Sheets.Add(After:= _
        wb.Sheets(wb.Sheets.Count))

    ws.Name = Left(LCase(sheetName), 16)
    
    'n is the count of numbers to be
    'chosen from the number set
    n = gd.Range("A2").Value

    'm is the number set size
    m = gd.Range("B2").Value
    
    For i = 1 To m
        z.Add CStr(i), CStr(i)
    Next
    
    Application.ScreenUpdating = False
    
    For i = 1 To WorksheetFunction.RoundUp(m / n, 0)
    
        Cells(i + 1, 1).Value = "Tipp " & i
        
        Set y = Nothing
        
        For t = 1 To n
               
            If z.Count < n And y.Count = 0 Then
            
                For c = 1 To m
                    y.Add CStr(c), CStr(c)
                Next
                
                For Each item In z
                    y.Remove (item)
                Next

            End If
            
            If z.Count = 0 Then
                Set z = Nothing
                Set z = y
            End If
            
            Randomize (Rnd * Now() + 1.0365 * 10)
            
            r = WorksheetFunction.RoundUp((Rnd * z.Count), 0)

            ws.Cells(i + 1, t + 1).Value = z.item(r)

            z.Remove (r)
            
            If t = n Then
                
                With ws.Sort
                    .SortFields.Clear
                    .SortFields.Add2 Key:=Range(Cells(i + 1, 2), _
                    Cells(i + 1, t + 1)), SortOn:=xlSortOnValues, _
                    Order:=xlAscending, DataOption:=xlSortTextAsNumbers
                    .SetRange Range(Cells(i + 1, 2), Cells(i + 1, t + 1))
                    .Header = xlNo
                    .MatchCase = False
                    .Orientation = xlLeftToRight
                    .SortMethod = xlPinYin
                    .Apply
                End With
        
            End If
            
            If i = WorksheetFunction.RoundUp(m / n, 0) And t = n Then
                            
                Set e = ws.Range(Cells(2, 2), Cells(i + 1, t + 1))
   
                e.FormatConditions.AddUniqueValues
                e.FormatConditions(e.FormatConditions.Count).SetFirstPriority
                
                With e.FormatConditions(1)
                    .DupeUnique = xlDuplicate
                    .Font.Color = -16383844
                    .Font.TintAndShade = 0
                    .Interior.PatternColorIndex = xlAutomatic
                    .Interior.Color = 13551615
                    .Interior.TintAndShade = 0
                End With
                
            End If
        
        Next
        
    DoEvents
    Next
    
    With ws
        .Cells.Interior.Color = vbWhite
        .Cells.Font.Name = "Tahoma"
        .Cells.Font.Size = "9"
        .Range("A:A").Cells.Font.Bold = True
        .Columns("A:Z").AutoFit
    End With
    
    gd.Activate
    
    Application.ScreenUpdating = True
    
    Set y = Nothing
    Set z = Nothing
    Set ws = Nothing
    Set gd = Nothing
    sheetName = vbNullString
   
End Sub




