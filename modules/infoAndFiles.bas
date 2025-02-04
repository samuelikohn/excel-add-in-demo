Attribute VB_Name = "infoAndFiles"
Sub exportFiles(Optional control As IRibbonControl)
    ''' Opens a file dialogue for the user to select a folder, then saves the tensile test template sheet as
    ''' an Excel workbook to the selected folder.
    
    
    
    Dim sourceWorkbook As Workbook
    Set sourceWorkbook = ActiveWorkbook
    Dim savePath As String
    Dim destinationWorkbook As Workbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Get save location
    savePath = Application.GetSaveAsFilename(FileFilter:=" Microsoft Excel Workbook (*.xlsx), *.xlsx", Title:="Select a Folder")
    If savePath = "False" Then GoTo exitHandler
    
    ' Create blank workbook
    Set destinationWorkbook = Workbooks.Add
    
    ' Copy template sheet to new workbook
    tensileTestDataSheet.Copy After:=destinationWorkbook.Sheets(1)
    
    ' Delete blank sheet
    destinationWorkbook.Sheets(1).Delete

    ' Save and close file
    destinationWorkbook.SaveAs fileName:=savePath, FileFormat:=xlWorkbookDefault
    destinationWorkbook.Close
    
exitHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub viewVideoGuide(Optional control As IRibbonControl)
    ''' Opens a YouTube video demonstrating the macros and forms in this add-in in your default web browser.
    
    
    
    ActiveWorkbook.FollowHyperlink Address:="https://www.youtube.com/watch?v=c0MCmpdpiKM"
    
End Sub

Sub viewInstallationGuide(Optional control As IRibbonControl)
    ''' Opens a PDF of the general Excel Add-in installation guide in your default PDF viewing program.



    embeddedFiles.OLEObjects("Object 1").Activate
    
End Sub

Sub viewAbout(Optional control As IRibbonControl)
    ''' Displays information about this Add-in.



    dummy = MsgBox("Excel Add-in Demo created by Sam Kohn." & vbCrLf & vbCrLf & "Github: https://github.com/samuelikohn/excel-add-in-demo" & vbCrLf & vbCrLf & "Contact: samuelikohn@gmail.com", Title:="About This Add-in")
    
End Sub
