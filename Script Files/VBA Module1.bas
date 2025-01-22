Attribute VB_Name = "Module1"
Sub WS_Loop()
Attribute WS_Loop.VB_ProcData.VB_Invoke_Func = " \n14"
'create a variable to hold the counter
Dim i, SummaryRow As Integer
'create variables
Dim ws As Worksheet
Dim OpenPrice, ClosePrice, QuarterlyChange, PercentageChange, totalVolume As Double
Dim lastRow, RowFirst As Long
'Variables to track greatest value
Dim GreatestIncreaseTicker, GreatestDecreaseTicker, GreatestVolumeTicker, Ticker As String
Dim GreatestIncrease, GreatestDecrease, GreatestVolume As Double

' Loop through each worksheet (quarter)
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
 ' Set up headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        SummaryRow = 2 ' Start the summary in row 2
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Loop through each stock entry in the worksheet
        For i = 2 To lastRow
        RowFirst = Columns(1).Find(What:=ws.Cells(i, 1).Value, Lookat:=xlWhole, SearchDirection:=xlNext, MatchCase:=False).Row
        ' Check if it's a new stock ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(RowFirst, 3).Value
            ClosePrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
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
                ws.Cells(SummaryRow, 12).Value = totalVolume
            ' Conditional formatting for quarterly change
                If QuarterlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive
                Else
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative
                End If
                ' Move to next row in the summary table and reset volume
                SummaryRow = SummaryRow + 1
                totalVolume = 0
            Else
                ' Accumulate volume if it's the same ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
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
    Next ws
End Sub
Sub Greatest()
Attribute Greatest.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ws As Worksheet
Dim MyMax, MyMin, MyMax2 As Variant
Dim maxCell, MinCell, MaxCell2 As Variant

For Each ws In ThisWorkbook.Worksheets

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Increase"

MyMax = Application.WorksheetFunction.Max(ws.Range("K:K"))
Set maxCell = ws.Range("K:K").Find(MyMax, Lookat:=xlWhole)

MyMin = Application.WorksheetFunction.Min(ws.Range("K:K"))
Set MinCell = ws.Range("K:K").Find(MyMin, Lookat:=xlWhole)

MyMax2 = Application.WorksheetFunction.Max(ws.Range("L:L"))
Set MaxCell2 = ws.Range("L:L").Find(MyMax2, Lookat:=xlWhole)

ws.Range("Q2").Value = MyMax
ws.Range("P2").Value = maxCell.Offset(, -2)

ws.Range("Q3").Value = MyMin
ws.Range("P3").Value = MinCell.Offset(, -2)

ws.Range("Q4") = MyMax2
ws.Range("P4") = MaxCell2.Offset(, -3)

Next ws
End Sub
