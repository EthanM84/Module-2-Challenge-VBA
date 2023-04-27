Attribute VB_Name = "Module1"
Sub StockAnalysis()

    Dim ws As Worksheet
    Dim years As Variant
    years = Array("2018", "2019", "2020") ' Array of years
    
    Dim year As Variant
    For Each year In years ' Loop through each year
        Set ws = ThisWorkbook.Sheets(year) ' Set the worksheet
        
        Dim LastRow As Long
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Find the last row
        
        Dim ticker As String
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        
        ' Set column headers in summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Dim summaryTableRow As Long
        summaryTableRow = 2 ' Start from row 2 in summary table
        
        Dim i As Long
        Dim firstRow As Long
        firstRow = 2 ' Start from row 2 in data table
        
        For i = 2 To LastRow ' Loop through all rows in data table
            
            ' Check if the ticker symbol has changed
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closingPrice = ws.Cells(i, 6).Value
                YearlyChange = closingPrice - ws.Cells(firstRow, 3).Value
                TotalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(firstRow, 7), ws.Cells(i, 7)))
                
                ' Output results to summary table
                ws.Cells(summaryTableRow, 9).Value = ticker
                ws.Cells(summaryTableRow, 10).Value = YearlyChange
                If ws.Cells(firstRow, 3).Value <> 0 Then ' Avoid division by zero
                    PercentChange = YearlyChange / ws.Cells(firstRow, 3).Value
                Else
                    PercentChange = 0
                End If
                ws.Cells(summaryTableRow, 11).Value = Format(PercentChange, "0.00%")
                ws.Cells(summaryTableRow, 12).Value = TotalVolume
                
                ' Format yearly change cell with conditional formatting
                If YearlyChange > 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf YearlyChange < 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Move to the next row in the summary table
                summaryTableRow = summaryTableRow + 1
                firstRow = i + 1 ' Update the first row for the next ticker symbol
            End If
            
        Next i
        
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        Dim LastSummaryRow As Long
                LastSummaryRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row ' Find the last row in the summary table
        
        Dim maxPercentIncrease As Double
        Dim maxPercentDecrease As Double
        Dim maxTotalVolume As Double
        Dim maxPercentIncreaseTicker As String
        Dim maxPercentDecreaseTicker As String
        Dim maxTotalVolumeTicker As String
        
        maxPercentIncrease = WorksheetFunction.Max(ws.Range("K2:K" & LastSummaryRow)) ' Find maximum percent increase
        maxPercentDecrease = WorksheetFunction.Min(ws.Range("K2:K" & LastSummaryRow)) ' Find maximum percent decrease
        maxTotalVolume = WorksheetFunction.Max(ws.Range("L2:L" & LastSummaryRow)) ' Find maximum total volume
        
        ' Find tickers for maximum percent increase, maximum percent decrease, and maximum total volume
        maxPercentIncreaseTicker = ws.Cells(WorksheetFunction.Match(maxPercentIncrease, ws.Range("K2:K" & LastSummaryRow), 0) + 1, 9).Value
        maxPercentDecreaseTicker = ws.Cells(WorksheetFunction.Match(maxPercentDecrease, ws.Range("K2:K" & LastSummaryRow), 0) + 1, 9).Value
        maxTotalVolumeTicker = ws.Cells(WorksheetFunction.Match(maxTotalVolume, ws.Range("L2:L" & LastSummaryRow), 0) + 1, 9).Value
        
        ' Output results to summary table
        ws.Cells(2, 17 - 1).Value = maxPercentIncreaseTicker
        ws.Cells(2, 18 - 1).Value = Format(maxPercentIncrease, "0.00%")
        ws.Cells(3, 17 - 1).Value = maxPercentDecreaseTicker
        ws.Cells(3, 18 - 1).Value = Format(maxPercentDecrease, "0.00%")
        ws.Cells(4, 17 - 1).Value = maxTotalVolumeTicker
        ws.Cells(4, 18 - 1).Value = maxTotalVolume
        
        ' Format percent increase and percent decrease cells as percentage
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0.00E+00"
        
        ' Auto-fit columns in summary table
        ws.Columns("I:Q").AutoFit
        
    Next year

End Sub


