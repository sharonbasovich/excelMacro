Attribute VB_Name = "Module1"
Sub GenerateApproximateValues()
'
' GenerateApproximateValues Macro
' Generates polynomial best fit line with specified degree and uses it to generate specified amount of values with specified step (increment between values)
'

'
' MsgBox Selection.Column
Degree = InputBox("Enter degree of polynomial estimate.", "Line of best fit")
step = InputBox("Enter the value to be incremented between data points", "Step")
' firstInCell = InputBox("Enter first cell of independent variable", "First Independent Cell")
startValue = InputBox("Enter first cell of independent variable (example: C6).", "First Independent Cell")
lastInCell = InputBox("Enter last cell of independent variable (example: C6).", "Last Independent Cell")
firstOutCell = InputBox("Enter first cell of dependent variable (example: C6).", "First Dependent Cell")
coeffLocation = InputBox("Enter cell for left most coefficient (subsequent coefficients will be displayed to the right).", "Coefficient Cell")
outputLocation1 = InputBox("Enter cell for first independent value output (subsequent values will be displayed below in a column).", "Independent Output")
outputLocation2 = InputBox("Enter cell for first dependent value output(subsequent values will be displayed below in a column).", "Dependent Output")
    Dim amount As Variant
    Dim startNum As Variant
    Dim endNum As Variant
    Dim terms As Variant
    Dim format As String
    Dim arr As Variant
    Dim tempTerm As Variant
    Dim coeff(1 To 100) As Variant
'   Dim startValue As Variant
    Dim endValue As Variant
    Dim extraNum As Boolean
    arr = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
' value of selected cell
    ActiveSheet.Range("" & startValue & "").Select
    startNum = Selection.Value
' value of last independent cell
    endNum = Range("" & lastInCell & "").Value
' amount of values to be generated
    amount = 1 + (endNum - startNum) / step
' amount of input values
    terms = Range(Selection.Address & ":" & "" & lastInCell & "").Cells.Count
    If (startNum + (step * amount)) > endNum Then
        extraNum = True
    End If
' get address of last dependent cell
    ActiveSheet.Range("" & firstOutCell & "").Select
    ActiveCell.Offset((terms - 1), 0).Range("A1").Select
    endValue = ActiveCell.Address(0, 0)
    ActiveSheet.Range("" & startValue & "").Select
' selects coeffecient cells
     ActiveSheet.Range("" & coeffLocation & "").Select
     ActiveCell.Range("A1:" & arr(Degree) & "1").Select
' sets format to "1, 2, 3, ... " up to degree
    format = 1
    For i = 2 To Degree
    format = format & ", " & i
    Next i
    tempTerm = terms
    Selection.FormulaArray = _
        "=LINEST(" & firstOutCell & ":" & endValue & "," & startValue & ":" & lastInCell & "^{" & format & "})"
    
    ActiveSheet.Range("" & outputLocation1 & "").Select
    ActiveCell.Value = "=" & startValue & ""
    ActiveCell.Offset(1, 0).Range("A1").Select
    For counter = 2 To amount
    ActiveCell.FormulaR1C1 = "=Sum(R[-1]C, " & step & ")"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Next counter
    ActiveSheet.Range("" & coeffLocation & "").Select
    For j = 1 To (Degree + 1)
    coeff(j) = Selection.Address(0, 0)
    ActiveCell.Offset(0, 1).Range("A1").Select
    Next j
    ActiveSheet.Range("" & outputLocation2 & "").Select
    Dim eval As String
    Dim temp As Variant
    Dim tempPos As Variant
    Dim tempCoeff As Variant
    For l = 1 To amount
        eval = "= IFERROR((0"
        tempPos = ActiveCell.Address(0, 0)
        ActiveSheet.Range("" & outputLocation1 & "").Select
        ActiveCell.Offset((l - 1), 0).Range("A1").Select
        temp = Selection.Address(0, 0)
        ActiveSheet.Range("" & tempPos & "").Select
        For k = 1 To (Degree + 1)
        tempCoeff = coeff(k)
            eval = eval & " + " & tempCoeff & " * " & temp & " ^ " & (Degree + 1 - k)
        Next k
        eval = eval & "), 0)"
        ActiveCell.Formula = "" & eval & ""
        ActiveCell.Offset(1, 0).Range("A1").Select
    Next l
End Sub




