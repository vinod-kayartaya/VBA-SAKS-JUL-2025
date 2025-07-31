Attribute VB_Name = "Module1"
Option Explicit

Sub CalculateSalesGrade()
    ' This subroutine calculates the Grade of
    ' sales based on the value of sales amount.
    ' If the sale amount is >= 1 lakh, then
    ' the grade is "Good" else the grade is "OK".
    ' This macro will read the value from the
    ' ActiveCell
    
    If IsEmpty(ActiveCell.Value) Then
        MsgBox "The active cell does not have a value to process." & vbCrLf & _
            "Select a cell that contains sales amount and then run this macro", _
            vbExclamation, "SAKS - Error"
        Exit Sub ' takes the control away from this subroutine, skipping the rest of the code
    End If
    
    If Not IsNumeric(ActiveCell.Value) Then
        MsgBox "Active cell contains a non-numeric value." & vbCrLf & _
            "Select a cell that contains sales amount and then run this macro", _
            vbExclamation, "SAKS - Error"
        Exit Sub
    End If
    
    Dim SalesAmount As Double
    Dim SalesGrade As String
    
    SalesAmount = ActiveCell.Value
    
    If SalesAmount >= 100000 Then
        SalesGrade = "Good"
    Else
        SalesGrade = "OK"
    End If
    
    ActiveCell.Offset(0, 1).Value = SalesGrade
    
End Sub


Sub FillSeries()
    ' This subroutine will fill down a series
    ' of values from 1 till N where N is
    ' accepted from the user. This will fill
    ' from the activecell downwards.
    
    Dim LowerLimit As Integer
    Dim UpperLimit As Integer, index% ' index as Integer
    UpperLimit = InputBox("Enter upper limit", "SAKS", 100)
    
    If IsEmpty(ActiveCell.Value) Or Not IsNumeric(ActiveCell.Value) Then
        LowerLimit = 1
    Else
        LowerLimit = ActiveCell.Value
    End If
    
    For index = LowerLimit To UpperLimit
        ActiveCell.Offset(index - LowerLimit).Value = index
    Next
    
End Sub


Sub CalculateAllGrades()
  
    If IsEmpty(ActiveCell) Then
        MsgBox "This can't be applied to the selected range. " & vbCrLf & _
            "Select a single cell from your data range, and try the same again", _
            vbExclamation, "SAKS"
        Exit Sub
    End If
    
    Dim DataRange As Range
    Dim c As Range
    Dim amount As Double
    
    If Selection.Cells.Count = 1 Then
        Set DataRange = ActiveCell.CurrentRegion
    Else
        Set DataRange = Selection
    End If

    Set DataRange = DataRange.Columns(2).Offset(1)
    Set DataRange = DataRange.Resize(DataRange.Rows.Count - 1)
    
    For Each c In DataRange.Cells
        amount = c.Value
        c.Offset(0, 1).Value = GetGradeForSalesAmount(amount)
    Next
End Sub

Function GetGradeForSalesAmount(amount As Double)

    Select Case amount
        Case 700000
            GetGradeForSalesAmount = "Awesome"
        Case Is >= 500000
            GetGradeForSalesAmount = "Excellent"
        Case Is >= 400000
            GetGradeForSalesAmount = "Very good"
        Case Is >= 300000
            GetGradeForSalesAmount = "Good"
        Case Is >= 150000
            GetGradeForSalesAmount = "Average"
        Case Else
            GetGradeForSalesAmount = "Not good"
    End Select

'        If amount >= 500000 Then
'            GetGradeForSalesAmount = "Excellent"
'        ElseIf amount >= 400000 Then
'            GetGradeForSalesAmount = "Very good"
'        ElseIf amount >= 300000 Then
'            GetGradeForSalesAmount = "Good"
'        ElseIf amount >= 150000 Then
'            GetGradeForSalesAmount = "Average"
'        Else
'            GetGradeForSalesAmount = "Not good"
'        End If
End Function

Sub HighlightBelowAverageValues()
    ' this macro will highlight those cells whose values are less than
    ' the average value of the selected data.
    
    Dim DataRange As Range
    
    ' check if the selection is a single cell or a range
    If Selection.Cells.Count > 1 Then
        Set DataRange = Selection
    Else
        Set DataRange = ActiveCell.CurrentRegion
    End If
    
    ' find the average of values in the DataRange
    Dim avg As Double
    avg = Application.WorksheetFunction.Average(DataRange)
    
    ' loop over all cells in the DataRange
    Dim c As Range
    For Each c In DataRange.Cells
        ' reset the formatting on c
        c.Font.Bold = False
        c.Font.Color = vbBlack
        c.Interior.Color = vbWhite
        
        ' check if the looped cell value is less than avg
        If c.Value < avg Then
            ' if yes, change the background and foreground color
            c.Font.Color = vbRed
            c.Interior.Color = vbYellow
            c.Font.Bold = True
        End If
    Next
    
End Sub




















