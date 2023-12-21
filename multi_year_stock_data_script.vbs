Sub stock_data_analysis()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    Dim IncreaseNumber As Long
    Dim DecreaseNumber As Long
    Dim VolumeNumber As Long
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row with data in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize summary row
        SummaryRow = 2
        
        ' Set headers for summary columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Perform analysis for each row in the worksheet
        For i = 2 To LastRow
            
            ' Retrieve data from the row
            Ticker = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(i, 3).Value
            ClosePrice = ws.Cells(i, 6).Value
            TotalVolume = ws.Cells(i, 7).Value
            
            ' Calculate Yearly Change and Percent Change
            If OpenPrice <> 0 Then
                YearlyChange = ClosePrice - OpenPrice
                PercentChange = (YearlyChange / OpenPrice) * 100
            Else
                YearlyChange = 0
                PercentChange = 0
            End If
            
            ' Store calculated values in summary columns
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            ws.Cells(SummaryRow, 11).Value = PercentChange
            ws.Cells(SummaryRow, 12).Value = TotalVolume
            
            ' Apply conditional formatting
            If YearlyChange > 0 Then
                ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf YearlyChange < 0 Then
                ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
            End If
            
            ' Move to the next summary row
            SummaryRow = SummaryRow + 1
            
        Next i
        
        ' Identify stocks with Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        ' (This part can be added based on your specific requirements)
        
    Next ws

End Sub
