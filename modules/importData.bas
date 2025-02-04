Attribute VB_Name = "importdata"
Sub importTensileTestCSV(Optional control As IRibbonControl)
    ''' Imports data from a formatted tensile test CSV file into a corresponding template sheet. If a
    ''' blank template sheet does not exist, one is created. The sheet is named according to the timestamp
    ''' of the tensile test. Series for stress, strain, and strain rate are calculated and a stress-strain
    ''' curve is plotted.
    


    Dim fileNames As Variant
    Dim fileName As Variant
    Dim destinationWorkbook As Workbook
    Set destinationWorkbook = ActiveWorkbook
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    Dim ws As Worksheet
    Dim destinationWorksheet As Worksheet
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim currentTime As Date: currentTime = Now()
    Dim AMPM As String
    Dim testTimeStamp As String
    Dim testHour As Integer
    Dim i As Long: i = 3
    Dim Area As Double
    Dim stressRange As Range
    Dim strainRange As Range
    Dim cht As Chart
    On Error GoTo exitHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Open file dialogue for selecting CSVs
    fileNames = Application.GetOpenFilename(FileFilter:="CSV Files (*.csv), *.csv", Title:="Select Tensile Test CSV Files", MultiSelect:=True)

    ' Iterate over all selected files
    For Each fileName In fileNames
    
        ' Open current source file
        Set sourceWorkbook = Workbooks.Open(fileName)
        Set sourceWorksheet = sourceWorkbook.Sheets(1)
        
        ' Search for empty tensile test template sheet
        For Each ws In destinationWorkbook.Worksheets
            If _
                Left(ws.Name, 11) = "TensileTest" And _
                ws.Cells(5, 3) = "" And _
                ws.Cells(6, 3) = "" And _
                ws.Cells(7, 3) = "" And _
                ws.Cells(10, 3) = "" And _
                ws.Cells(11, 3) = "" And _
                ws.Cells(12, 3) = "" And _
                ws.Cells(15, 2) = "" And _
                ws.Cells(15, 3) = "" And _
                ws.Cells(15, 4) = "" And _
                ws.Cells(15, 6) = "" And _
                ws.Cells(15, 7) = "" And _
                ws.Cells(15, 8) = "" _
            Then
                Set destinationWorksheet = ws
                Exit For
            End If
        Next
        
        ' If no blank sheet found, import one
        If destinationWorksheet Is Nothing Then
            Set ws = tensileTestDataSheet
            
            ' Import selected sheet after current sheet
            ws.Copy After:=currentSheet
            
            ' Set name based on current timestamp
            Set destinationWorksheet = destinationWorkbook.Sheets("Template Tensile Test")
            destinationWorksheet.Name = _
                "TensileTest " & DatePart("yyyy", currentTime) & _
                "-" & DatePart("m", currentTime) & _
                "-" & DatePart("d", currentTime) & _
                " " & DatePart("h", currentTime) & _
                ";" & DatePart("n", currentTime) & _
                ";" & DatePart("s", currentTime)
            
            'If current sheet is blank, delete when template sheet is imported
            If WorksheetFunction.CountA(currentSheet.UsedRange) = 0 And currentSheet.Shapes.Count = 0 Then
                currentSheet.Delete
            End If
            
        End If
        
        ' User ID
        destinationWorksheet.Cells(5, 3) = sourceWorksheet.Cells(1, 6)
        
        ' Test Date
        destinationWorksheet.Cells(6, 3) = Mid(sourceWorksheet.Cells(2, 6), 6, 5) & "-" & Left(sourceWorksheet.Cells(2, 6), 4)
        
        ' Test Time
        testTimeStamp = sourceWorksheet.Cells(2, 6)
        testHour = Mid(testTimeStamp, 12, 2)
        If testHour < 12 Then
            AMPM = " AM"
        Else
            AMPM = " PM"
            testHour = testHour - 12
        End If
        destinationWorksheet.Cells(7, 3) = testHour & ":" & Mid(testTimeStamp, 15, 2) & ":" & Mid(testTimeStamp, 18, 2) & AMPM
        
        ' Rename worksheet based on test timestamp
        destinationWorksheet.Name = _
            "TensileTest " & Mid(testTimeStamp, 1, 4) & _
            "-" & Mid(testTimeStamp, 6, 2) & _
            "-" & Mid(testTimeStamp, 9, 2) & _
            " " & Mid(testTimeStamp, 12, 2) & _
            ";" & Mid(testTimeStamp, 15, 2) & _
            ";" & Mid(testTimeStamp, 18, 2)
        
        ' Sample Length
        destinationWorksheet.Cells(10, 3) = sourceWorksheet.Cells(5, 6)
        
        ' Sample Width
        destinationWorksheet.Cells(11, 3) = sourceWorksheet.Cells(6, 6)
        
        ' Sample Thickness
        destinationWorksheet.Cells(12, 3) = sourceWorksheet.Cells(7, 6)
        
        ' Test Data
        While sourceWorksheet.Cells(i, 1) <> ""
            
            ' Force
            destinationWorksheet.Cells(i + 12, 2) = sourceWorksheet.Cells(i, 3)
            
            ' Extension
            destinationWorksheet.Cells(i + 12, 3) = sourceWorksheet.Cells(i, 2)
            
            ' Time
            destinationWorksheet.Cells(i + 12, 4) = sourceWorksheet.Cells(i, 1)
            
            ' Stress
            Area = sourceWorksheet.Cells(6, 6) * sourceWorksheet.Cells(7, 6)
            destinationWorksheet.Cells(i + 12, 6) = sourceWorksheet.Cells(i, 3) / Area
            
            ' Strain
            destinationWorksheet.Cells(i + 12, 7) = sourceWorksheet.Cells(i, 2) / sourceWorksheet.Cells(5, 6)
            
            ' Strain Rate
            If i > 3 Then
                destinationWorksheet.Cells(i + 12, 8) = (destinationWorksheet.Cells(i + 12, 7) - destinationWorksheet.Cells(i + 11, 7)) / (destinationWorksheet.Cells(i + 12, 4) - destinationWorksheet.Cells(i + 11, 4))
            End If
            
            i = i + 1
            
        Wend
        
        ' Close source file
        sourceWorkbook.Close
        
        ' Create chart
        Set stressRange = destinationWorksheet.Range("F15:F" & i + 11)
        Set strainRange = destinationWorksheet.Range("G15:G" & i + 11)
        Set cht = destinationWorksheet.Shapes.AddChart2(240, xlXYScatterLines, destinationWorksheet.Cells(2, 11).Left, destinationWorksheet.Cells(2, 11).Top, 500, 400).Chart
    
        ' Set data to chart
        cht.FullSeriesCollection(1).Name = ""
        cht.FullSeriesCollection(1).XValues = strainRange
        cht.FullSeriesCollection(1).Values = stressRange
        
        ' Title and axis labels
        cht.ChartTitle.Text = "Stress vs. Strain Curve"
        
        With cht.Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Strain (mm/mm)"
        End With
        
        With cht.Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Stress (MPa)"
        End With
        
    Next
    
exitHandler:
    Debug.Print Err.Description
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub importEmployeeJSON(Optional control As IRibbonControl)
    ''' Imports data from an employees JSON file into a corresponding template sheet. If a template sheet
    ''' does not already exist, one is created. A new pair of columns is added to the template sheet for
    ''' each JSON file.
    
    
    
    Dim fileNames As Variant
    Dim fileName As Variant
    Dim destinationWorkbook As Workbook
    Set destinationWorkbook = ActiveWorkbook
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    Dim ws As Worksheet
    Dim destinationWorksheet As Worksheet
    Dim JSON As String
    Dim employeeData As Object
    Dim columnIndex As Long
    Dim i As Integer
    Dim key As Variant
    On Error GoTo exitHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Open file dialogue for selecting JSON files
    fileNames = Application.GetOpenFilename(FileFilter:="JSON Files (*.json), *json", Title:="Select Employees JSON Files", MultiSelect:=True)
    
    ' Search for empty employees template sheet
    For Each ws In destinationWorkbook.Worksheets
        If ws.CodeName = "employeesDataSheet" Then
            Set destinationWorksheet = ws
            Exit For
        End If
    Next
        
    ' If no blank sheet found, import one
    If destinationWorksheet Is Nothing Then
        Set ws = employeesDataSheet
        
        ' Import selected sheet after current sheet
        ws.Copy After:=currentSheet
        
        ' Set name to "Employees"
        Set destinationWorksheet = destinationWorkbook.Sheets("Template Employees")
        destinationWorksheet.Name = "Employees"
        
        'If current sheet is blank, delete when template sheet is imported
        If WorksheetFunction.CountA(currentSheet.UsedRange) = 0 And currentSheet.Shapes.Count = 0 Then
            currentSheet.Delete
        End If
        
    End If
    
    ' Iterate over all selected files
    For Each fileName In fileNames
    
        'Import the contents of a JSON file as a text stream
        With CreateObject("ADODB.Stream")
            .Charset = "utf-8"
            .Open
            .LoadFromFile fileName
            JSON = .ReadText
        End With

        'Use JSON parser to create dictionary of JSON paths and values
        Set employeeData = jsonParser.parseJSON(JSON)
    
        ' Find first empty column
        If WorksheetFunction.CountA(destinationWorksheet.UsedRange) = 0 Then
            columnIndex = 1
        Else
            columnIndex = destinationWorksheet.UsedRange.Columns.Count + 1
        End If

        ' Add right border
        With destinationWorksheet.Columns(columnIndex + 1).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        ' Merge, center, and bold company name
        With destinationWorksheet.Range(Cells(1, columnIndex), Cells(1, columnIndex + 1))
            .HorizontalAlignment = xlCenter
            .Merge
            .Font.Bold = True
        End With
        
        ' Add bottom border and center "Name" and "Email" titles
        With destinationWorksheet.Range(Cells(2, columnIndex), Cells(2, columnIndex + 1))
            .HorizontalAlignment = xlCenter
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        
        ' Populate company name
        destinationWorksheet.Cells(1, columnIndex) = employeeData("root.name")
        
        ' Populate "Name" and "Email" titles
        destinationWorksheet.Cells(2, columnIndex) = "Name"
        destinationWorksheet.Cells(2, columnIndex + 1) = "Email"
        
        ' Populate employee data
        i = 0
        key = "root.employees[" & i & "]"
        While employeeData.Exists(key & ".name")
            destinationWorksheet.Cells(i + 3, columnIndex) = employeeData(key & ".name")
            destinationWorksheet.Cells(i + 3, columnIndex + 1) = employeeData(key & ".email")
            i = i + 1
            key = "root.employees[" & i & "]"
        Wend
        
        ' Autofit columns
        destinationWorksheet.Columns(columnIndex).autoFit
        destinationWorksheet.Columns(columnIndex + 1).autoFit
        
    Next
    
exitHandler:
    Debug.Print Err.Description
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub addEmployeesManually(Optional control As IRibbonControl)
    ''' Displays a user form for manually adding employee information to a data sheet.
    
    
    
    Dim employeesForm As New addEmployeesUserForm

    employeesForm.Show
    
End Sub

Sub importTemplateSheets(Optional control As IRibbonControl)
    ''' Displays a template sheets user form.



    Dim templateSheetsForm As New importTemplateSheetsUserForm

    templateSheetsForm.Show
    
End Sub

