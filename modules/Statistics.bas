Attribute VB_Name = "Statistics"
Sub descriptiveStatistics(Optional control As IRibbonControl)
    ''' Calculates descriptive statistics for the values in a selected range of cells. The following are
    ''' calculated for all values, regardless of data type:
    '''     Total number of values
    '''     Number of unique values
    '''     Mode (multiple values will be listed if applicable)
    '''     Mode Frequency
    '''
    ''' The following are calculated using only numerical values in the selected range:
    '''     Number of numerical values
    '''     Mean
    '''     Standard deviation
    '''     Minimum
    '''     1st quartil
    '''     Median
    '''     3rd quartile
    '''     Maximum
    '''     Range
    '''
    ''' Results are displayed in a user form which allows for copying to the clipboard.
    
    
    
    Dim sourceRange As Range
    Dim Cell As Range
    Dim Freq As Object
    Set Freq = CreateObject("Scripting.Dictionary")
    Dim Count As Long: Count = 0
    Dim maxFreq As Long: maxFreq = 0
    Dim Mode As String: Mode = ""
    Dim key As Variant
    Dim numberTotal As Double: numberTotal = 0
    Dim numberCount As Long: numberCount = 0
    Dim Result As String
    Dim resultsForm As descriptiveStatisticsUserForm
    Set resultsForm = New descriptiveStatisticsUserForm
    Dim Mean
    Dim standardDeviation
    Dim firstQuartile
    Dim Median
    Dim thirdQuartile
    On Error Resume Next
    
    ' Get range of data from user input
    Set sourceRange = Application.InputBox(Prompt:="Select a range of data to describe.", Title:="Descriptive Statistics", Type:=8)
    If Err.Number <> 0 Then
        Exit Sub ' User canceled
    End If

    ' Create frequency dictionary
    For Each Cell In sourceRange.Cells
    
        ' Increment total count
        Count = Count + 1
        
        ' Update frequency dictionary
        If Freq.Exists(Cell.Value) Then
            Freq(Cell.Value) = Freq(Cell.Value) + 1
        Else:
            Freq.Add Cell.Value, 1
        End If
        
        ' Keep track of most frequent value
        maxFreq = IIf(Freq(Cell.Value) > maxFreq, Freq(Cell.Value), maxFreq)
        
        ' Filter only numerical values
        If IsNumeric(Cell.Value) Then
            numberTotal = numberTotal + Cell.Value
            numberCount = numberCount + 1
        End If
        
    Next
    
    ' Find key with highest frequency value
    For Each key In Freq.keys()
        If Freq(key) = maxFreq Then
            Mode = Mode & key & ", "
        End If
    Next
    
    ' Handle case for worksheet functions when not enough numerical values are selected.
    If numberCount = 0 Then
        Mean = ""
        standardDeviation = ""
        firstQuartile = ""
        Median = ""
        thirdQuartile = ""
        
    ElseIf numberCount = 1 Then
        standardDeviation = ""
        firstQuartile = WorksheetFunction.Quartile(sourceRange, 1)
        Median = WorksheetFunction.Median(sourceRange)
        thirdQuartile = WorksheetFunction.Quartile(sourceRange, 3)
        Mean = numberTotal / numberCount
        
    Else
        standardDeviation = WorksheetFunction.StDev(sourceRange)
        firstQuartile = WorksheetFunction.Quartile(sourceRange, 1)
        Median = WorksheetFunction.Median(sourceRange)
        thirdQuartile = WorksheetFunction.Quartile(sourceRange, 3)
        Mean = numberTotal / numberCount
        
    End If
    
    ' Display results
    Result = _
        "Total Count:  " & vbTab & Count & vbCrLf & _
        "Unique Count:" & vbTab & Freq.Count & vbCrLf & _
        "Mode(s):       " & vbTab & Left(Mode, Len(Mode) - 2) & vbCrLf & _
        "Mode Count: " & vbTab & maxFreq & vbCrLf & vbCrLf & _
        "Number Count:" & vbTab & numberCount & vbCrLf & _
        "Mean:           " & vbTab & Mean & vbCrLf & _
        "Standard Deviation:" & vbTab & standardDeviation & vbCrLf & _
        "Minimum:     " & vbTab & WorksheetFunction.Min(sourceRange) & vbCrLf & _
        "1st Quartile:  " & vbTab & firstQuartile & vbCrLf & _
        "Median:        " & vbTab & Median & vbCrLf & _
        "3rd Quartile: " & vbTab & thirdQuartile & vbCrLf & _
        "Maximum:     " & vbTab & WorksheetFunction.Max(sourceRange) & vbCrLf & _
        "Range:          " & vbTab & WorksheetFunction.Max(sourceRange) - WorksheetFunction.Min(sourceRange)
        
    resultsForm.statisticsTextBox.MultiLine = True
    resultsForm.statisticsTextBox.Value = Result
    resultsForm.Show

End Sub

