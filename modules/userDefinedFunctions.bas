Attribute VB_Name = "userDefinedFunctions"
Public Function CORRELMATRIX(inputRange As Range) As Variant
    ''' Calculates a correlation matrix for series of data using Excel's "CORREL" worksheet function.
    '''
    ''' Args:
    '''     inputRange (Range): Contiguous range of cells containing the data series. Each column represents a single series.
    '''
    ''' Returns:
    '''     N by N matrix of numbers between 0 and 1 representing correlation coefficients, where N is the number of series in InputRange.
    '''
    ''' Errors:
    '''     Raises a #VALUE! error if any cells in InputRange contain non-numerical values.
    '''     Raises a #VALUE! error if any data series only contains 1 value.
    '''     Raises a #SPILL! error if the returned matrix would overflow into occupied cells.
    '''     Raises a #VALUE! error if the CORREL function returns an error.


    Dim numCols As Long
    Dim correlationArray() As Double
    On Error GoTo errorHandler

    ' Check for single-cell range or non-numeric data
    If inputRange.Cells.Count = 1 Or Not IsNumeric(inputRange.Cells(1, 1).Value) Then
        Err.Raise Number:=vbObjectError, Description:="Invalid input range. Please select a range with two or more columns of numerical data."
    End If

    ' Get the number of columns in the input range.
    numCols = inputRange.Columns.Count

    ' Create a 2D array to store the correlation results.
    ReDim correlationArray(1 To numCols, 1 To numCols)
 
    ' Calculate the correlation for each pair of columns.
    Dim i As Long, j As Long
    For i = 1 To numCols
        For j = 1 To numCols
        
            ' Check if both columns contain numerical data.
            If Application.WorksheetFunction.Count(inputRange.Columns(i)) <> inputRange.Rows.Count Or _
               Application.WorksheetFunction.Count(inputRange.Columns(j)) <> inputRange.Rows.Count Then
                Err.Raise Number:=vbObjectError, Description:="Invalid input range: One or more columns contain non-numeric data."
            End If
             
            'Calculate the correlation between the current and j-th columns.
            correlationArray(i, j) = Application.WorksheetFunction.Correl(inputRange.Columns(i), inputRange.Columns(j))
        Next j
    Next i

    ' Transpose the CorrelationArray to display it as a matrix.
    CORRELMATRIX = Application.WorksheetFunction.Transpose(correlationArray)

    Exit Function

errorHandler:

    ' Return #VALUE! error if an error occurs.
    CORRELMATRIX = CVErr(xlErrValue)
    Debug.Print "Error " & Err.Number & ": " & Err.Description

End Function

Sub viewUDFDocumentation(Optional control As IRibbonControl)
    ''' Displays a "View UDF Documentation" user form.



    Dim viewUDFForm As New viewUDFsUserForm

    viewUDFForm.Show

End Sub
