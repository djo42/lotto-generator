Attribute VB_Name = "prc_lotto"
Option Explicit

Sub lotto()

    Dim sheetName As String
    Dim namePart As String
    Dim wb As Workbook
    Dim ws As Variant
    Dim e As Range
    Dim n As Long
    Dim m As Long
    Dim t As Long
    Dim i As Long
    Dim y As New Collection
    Dim z As New Collection
    Dim r As Long
    Dim c As Long
    Dim item As Variant
     
    Set wb = ThisWorkbook
    sheetName = GUID(True, False)

    Set ws = wb.Sheets.Add(After:= _
        wb.Sheets(wb.Sheets.Count))

    ws.Name = Left(LCase(sheetName), 16)
    
    n = wb.Sheets("GameData").Range("A2").Value
    m = wb.Sheets("GameData").Range("B2").Value
    
    For i = 1 To m
    
        With z
            .Add CStr(i), CStr(i)
        End With
    
    Next
    
    Application.ScreenUpdating = False
    
    For i = 1 To WorksheetFunction.RoundUp(m / n, 0)
    
        Cells(i + 1, 1).Value = "Tipp " & i
        
        Set y = Nothing
        
        For t = 1 To n
               
            If z.Count < n And y.Count = 0 Then
            
                For c = 1 To m
                    With y
                        .Add CStr(c), CStr(c)
                    End With
                Next
                
                For Each item In z
                    With y
                        y.Remove (item)
                    End With
                Next

            End If
            
            If z.Count = 0 Then
                Set z = Nothing
                Set z = y
                                
            End If
            
            Debug.Print (Rnd * Now())
            Randomize (Rnd * Now())
            
            r = WorksheetFunction.RoundUp((Rnd * z.Count), 0)

            With ws
                .Cells(i + 1, t + 1).Value = z.item(r)
                
            End With
            z.Remove (r)
            
            If t = n Then
               
                ws.Sort.SortFields.Clear
                ws.Sort.SortFields.Add2 Key:=Range(Cells(i + 1, 2), Cells(i + 1, t + 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
                
                With ws.Sort
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
                e.FormatConditions(1).DupeUnique = xlDuplicate
                With e.FormatConditions(1).Font
                    .Color = -16383844
                    .TintAndShade = 0
                End With
                With e.FormatConditions(1).Interior
                    .PatternColorIndex = xlAutomatic
                    .Color = 13551615
                    .TintAndShade = 0
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
    
    Worksheets("GameData").Activate
    
    Application.ScreenUpdating = True
    
    Set y = Nothing
    Set z = Nothing
   
End Sub


Function GUID$(Optional lowercase As Boolean, Optional parens As Boolean)
    Dim k&, h$
    GUID = Space(36)
    For k = 1 To Len(GUID)
        Randomize
        Select Case k
            Case 9, 14, 19, 24: h = "-"
            Case 15:            h = "4"
            Case 20:            h = Hex(Rnd * 3 + 8)
            Case Else:          h = Hex(Rnd * 15)
        End Select
        Mid$(GUID, k, 1) = h
    Next
    If lowercase Then GUID = LCase$(GUID)
    If parens Then GUID = "{" & GUID & "}"
End Function























