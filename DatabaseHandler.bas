Attribute VB_Name = "Module1"
Sub Macro2()
'
' Macro4 Macro
Dim sheetCount As Integer
Dim counter As Integer
sheetCount = ThisWorkbook.Sheets.Count
Sheets(2).Activate
Range("A2").Select
Dim subCount As Integer
subCount = 0

Do Until IsEmpty(ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    subCount = subCount + 1
Loop
    
Dim subNames() As String
ReDim subNames(1 To subCount)
Range("A2").Select

For k = 1 To subCount
    subNames(k) = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
Next k
Dim valid As Integer
valid = 0
For i = 4 To sheetCount
    Sheets(i).Activate
                Columns("G:G").ColumnWidth = 14
                Range("G1").Select
                Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                   Formula1:="=0"
                Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                With Selection.FormatConditions(1).Font
                   .Color = -16754788
                   .TintAndShade = 0
                End With
                With Selection.FormatConditions(1).Interior
                   .PatternColorIndex = xlAutomatic
                   .Color = 10284031
                   .TintAndShade = 0
                End With
                Selection.FormatConditions(1).StopIfTrue = False
                With Selection
                   .HorizontalAlignment = xlLeft
                   .VerticalAlignment = xlBottom
                   .WrapText = False
                   .Orientation = 0
                   .AddIndent = False
                   .IndentLevel = 0
                   .ShrinkToFit = False
                   .ReadingOrder = xlContext
                   .MergeCells = False
                End With
                Selection.FormatConditions.Add Type:=xlTextString, String:="Not found", _
                   TextOperator:=xlContains
                Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                With Selection.FormatConditions(1).Font
                   .Color = -16383844
                   .TintAndShade = 0
                End With
                With Selection.FormatConditions(1).Interior
                   .PatternColorIndex = xlAutomatic
                   .Color = 13551615
                   .TintAndShade = 0
                End With
                Selection.FormatConditions(1).StopIfTrue = False
                Selection.FormatConditions.Add Type:=xlTextString, String:=",", _
                   TextOperator:=xlContains
                Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                With Selection.FormatConditions(1).Font
                   .Color = -16752384
                   .TintAndShade = 0
                End With
                With Selection.FormatConditions(1).Interior
                   .PatternColorIndex = xlAutomatic
                   .Color = 13561798
                   .TintAndShade = 0
                End With
                Selection.FormatConditions(1).StopIfTrue = False
                ActiveCell.FormulaR1C1 = _
                   "=XLOOKUP(RC[-4],'Part Numbers'!C[-5],'Part Numbers'!C[-6],""Not Found"")"
               
                Range("A1").Select
                counter = 0
                Do Until IsEmpty(ActiveCell.Value)
                   ActiveCell.Offset(1, 0).Select
                   counter = counter + 1
                Loop
                Range("G1").Select
                Selection.AutoFill Destination:=Range("G1:G" & "" & counter & ""), Type:=xlFillDefault
                Dim currentCell As String
                Dim isTeamcenter As Boolean
                
                For l = 1 To counter
                    Range("A" & l).Select
                    isTeamcenter = False
                    currentCell = ActiveCell.Value
                    For j = 1 To subCount
                        If InStr(currentCell, subNames(j)) > 0 Then
                            isTeamcenter = True
                            j = subCount
                        End If
                    Next j
                    Range("H" & l).Select
                    If isTeamcenter Then
                        ActiveCell.Value = "Found"
                    Else
                        ActiveCell.Value = "Not Found"
                    End If
                Next l
        Dim numFound As Boolean
        Dim isLibray As Boolean
        Dim isMagna As Boolean
        
        For m = 1 To counter
        Range("G" & m).Select
        numFound = False
        isLibrary = False
        isMagna = False
        If ActiveCell.Value = "Found" Then numFound = True
        Range("B" & m).Select
        If ActiveCell.Value = "LIBRARY" Then isLibrary = True
        Range("H" & m).Select
        If ActiveCell.Value = "Found" Then isMagna = True
        If (Not (numFound) And isMagna And Not (isLibrary)) Then
            valid = valid + 1
            Range("A" & m & ":C" & m).Select
            Selection.Copy
            Sheets(3).Activate
            Range("A" & valid).Select
            ActiveSheet.Paste
            Sheets(i).Activate
        End If
    Next m
Next i
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'
    Dim i As Integer
    i = 4
    Sheets(i).Activate
    
End Sub
