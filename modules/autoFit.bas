Attribute VB_Name = "autoFit"
Sub autoFitAllColumns(Optional control As IRibbonControl)
    ''' Adjusts the width of all columns on the current spreadsheet to automatically fit their contents.
    
    
    
    ActiveSheet.Cells.EntireColumn.autoFit
    
End Sub

Sub autoFitAllRows(Optional control As IRibbonControl)
    ''' Adjusts the height of all rows on the current spreadsheet to automatically fit their contents.
    
    
    
    ActiveSheet.Cells.EntireRow.autoFit
    
End Sub

Sub resetColumnWidth(Optional control As IRibbonControl)
    ''' Resets the width of all columns on the current spreadsheet to their standard width.



    ActiveSheet.Cells.EntireColumn.UseStandardWidth = True
    
End Sub

Sub resetRowHeight(Optional control As IRibbonControl)
    ''' Resets the height of all rows on the current spreadsheet to their standard height.
    
    
    
    ActiveSheet.Cells.EntireRow.UseStandardHeight = True

End Sub

