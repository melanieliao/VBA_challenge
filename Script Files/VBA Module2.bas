Attribute VB_Name = "Module2"
Sub WS_Loop2()
Attribute WS_Loop2.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim startOpenPrice As Double, endClosePrice As Double
    Dim totalVolume As Double
    Dim i As Long
    Dim change As Double
    Dim percentChange As Double
    Dim outputRow As Long
    Dim sheetIndex As Integer
    
    ' Start output from row 2 (skip header row) in the first sheet
    outputRow = 2
    
    ' Loop through the 4 sheets (representing each quarter)
    For sheetIndex = 1 To 4
        Set ws = ThisWorkbook.Sheets(sheetIndex)  ' Change to the correct sheet (Quarter 1, 2, 3, 4)
        
        ' Get the last row of data in column A for the current sheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through all the rows in the current sheet (start from row 2)
        i = 2
        Do While i <= lastRow
            
            ' Get the ticker symbol for the current stock
            tickerSymbol = ws.Cells(i, 1).Value
            
            ' Initialize variables for the current ticker's quarter
            startOpenPrice = 0
            endClosePrice = 0
            totalVolume = 0
            
            ' Process rows for the same ticker symbol (same quarter)
            Do While i <= lastRow And ws.Cells(i, 1).Value = tickerSymbol
                
                ' For the first row of the ticker symbol (start of quarter), capture the opening price
                If startOpenPrice = 0 Then
                    startOpenPrice = ws.Cells(i, 3).Value  ' Column C: Opening price
                End If
                
                ' Capture the closing price (last row for this ticker symbol in the quarter)
                endClosePrice = ws.Cells(i, 6).Value  ' Column F: Closing price
                
                ' Sum the volume for all rows of the same ticker symbol in the quarter
                totalVolume = totalVolume + ws.Cells(i, 7).Value  ' Column G: Volume
                
                ' Move to the next row
                i = i + 1
                
            Loop
            
            ' Calculate the change and percentage change for the quarter
            If startOpenPrice <> 0 And endClosePrice <> 0 Then
                change = endClosePrice - startOpenPrice
                percentChange = (change / startOpenPrice) * 100
                
                ' Output results for the current ticker's quarter
                ws.Cells(outputRow, 9).Value = tickerSymbol  ' Ticker symbol in Column H
                ws.Cells(outputRow, 10).Value = change  ' Quarterly change in Column I
                ws.Cells(outputRow, 11).Value = percentChange  ' Percentage change in Column J
                ws.Cells(outputRow, 12).Value = totalVolume  ' Total volume in Column K
                
                ' Increment the output row for the next quarter
                outputRow = outputRow + 1
            End If
            
        Loop
        
    Next sheetIndex
    MsgBox "Quarterly Analysis Completed!"
End Sub
Sub WS_Loop3()
Attribute WS_Loop3.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim startOpenPrice As Double, endClosePrice As Double
    Dim totalVolume As Double
    Dim i As Long
    Dim change As Double
    Dim percentChange As Double
    Dim outputRow As Long
    Dim sheetNames As Variant
    Dim sheetIndex As Integer
    
    ' Sheet names for quarters
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    ' Loop through the 4 sheets (representing each quarter)
    For sheetIndex = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(sheetIndex))  ' Access sheet by name
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Get the last row of data in column A for the current sheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Start output from row 2 (skip header row) in each sheet
        outputRow = 2
        
        ' Loop through all the rows in the current sheet (start from row 2)
        i = 2
        Do While i <= lastRow
            
            ' Get the ticker symbol for the current stock
            tickerSymbol = ws.Cells(i, 1).Value
            
            ' Initialize variables for the current ticker's quarter
            startOpenPrice = 0
            endClosePrice = 0
            totalVolume = 0
            
            ' Process rows for the same ticker symbol (same quarter)
            Do While i <= lastRow And ws.Cells(i, 1).Value = tickerSymbol
                
                ' For the first row of the ticker symbol (start of quarter), capture the opening price
                If startOpenPrice = 0 Then
                    startOpenPrice = ws.Cells(i, 3).Value  ' Column C: Opening price
                End If
                
                ' Capture the closing price (last row for this ticker symbol in the quarter)
                endClosePrice = ws.Cells(i, 6).Value  ' Column F: Closing price
                
                ' Sum the volume for all rows of the same ticker symbol in the quarter
                totalVolume = totalVolume + ws.Cells(i, 7).Value  ' Column G: Volume
                
                ' Move to the next row
                i = i + 1
                
            Loop
            
            ' Calculate the change and percentage change for the quarter
            If startOpenPrice <> 0 And endClosePrice <> 0 Then
                change = endClosePrice - startOpenPrice
                percentChange = (change / startOpenPrice) * 100
                
                ' Output results for the current ticker's quarter
                ws.Cells(outputRow, 9).Value = tickerSymbol  ' Ticker symbol in Column H
                ws.Cells(outputRow, 10).Value = change  ' Quarterly change in Column I
                ws.Cells(outputRow, 11).Value = percentChange  ' Percentage change in Column J
                ws.Cells(outputRow, 12).Value = totalVolume  ' Total volume in Column K
                
                ' Increment the output row for the next ticker's data
                outputRow = outputRow + 1
            End If
            
        Loop
        
        With ws.Range("K2:K" & lastRow)
            .FormatConditions.Delete ' Clear existing formatting
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative
        End With

        
    Next sheetIndex
    
    MsgBox "Quarterly Analysis Completed!"


End Sub
Sub Greatest2()
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

