Attribute VB_Name = "Module1"
Sub alphabetical_testing()
'create a variable to hold the counter
Dim i As Integer

'create variables
Dim ws As Worksheet
Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim QuarterlyChange As Double
Dim PercentageChange As Double
Dim TotalVolume As Double
Dim LastRow As Long
Dim SummaryRow As Integer

'Variables to track greatest value
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestVolumeTicker As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double
    
'Initialize Summary values
GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

' Loop through each worksheet (quarter)
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
 ' Set up headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        SummaryRow = 2 ' Start the summary in row 2
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through each stock entry in the worksheet
        For i = 2 To LastRow
        
         OpenPrice = ws.Cells(i, 3).Value
        ' Check if it's a new stock ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ClosePrice = ws.Cells(i, 6).Value
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
        ' Calculate quarterly change and percentage change
        QuarterlyChange = ClosePrice - OpenPrice
            If OpenPrice <> 0 Then
            PercentageChange = (QuarterlyChange / OpenPrice)
            Else
            PercentageChange = 0
            End If
                
            ' Output the data to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = QuarterlyChange
                ws.Cells(SummaryRow, 11).Value = PercentageChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
            ' Conditional formatting for quarterly change
                If QuarterlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive
                Else
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative
                End If
                
            ' Check for greatest increase, decrease, and volume
                If PercentageChange > GreatestIncrease Then
                    GreatestIncrease = PercentageChange
                    GreatestIncreaseTicker = Ticker
                End If
                
                If PercentageChange < GreatestDecrease Then
                    GreatestDecrease = PercentageChange
                    GreatestDecreaseTicker = Ticker
                End If
                
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If
                
                ' Move to next row in the summary table and reset volume
                SummaryRow = SummaryRow + 1
                TotalVolume = 0
                
            Else
                ' Accumulate volume if it’s the same ticker
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Add summary for greatest values at the end of each worksheet
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = GreatestIncreaseTicker
        ws.Cells(2, 17).Value = GreatestIncrease

        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = GreatestDecreaseTicker
        ws.Cells(3, 17).Value = GreatestDecrease

        ws.Cells(4, 15).Value = "Greatest Volume"
        ws.Cells(4, 16).Value = GreatestVolumeTicker
        ws.Cells(4, 17).Value = GreatestVolume
        
        ws.Cells(1, 16).Value = "ticker"
        ws.Cells(1, 17).Value = "Value"
    Next ws

End Sub
 
