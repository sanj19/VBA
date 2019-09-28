Sub Stocks()

'Set variables
    Dim Symbol As String
    Dim StockTotal As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim SummaryTableRow As Integer
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim WorksheetName As String
    Dim ws As Worksheet
    Dim ws_num As Integer
    
'Get the number of worksheets in the workbook
ws_num = ThisWorkbook.Worksheets.Count

 'Loop through each worksheet
For a = 1 To ws_num
ThisWorkbook.Worksheets(a).Activate

    SummaryTableRow = 2
    
' Find the last row of the table
    lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

'Sort data by Col A, then Col B. This keeps all stock symbols grouped together and ensures
'     the first row will be Start of Year and last row will be End of Year (for each stock symbol)
    ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), Order:=xlAscending
    ActiveSheet.Sort.SortFields.Add Key:=Range("B1"), Order:=xlAscending
    ActiveSheet.Sort.SetRange Range("A1", Range("G1").End(xlDown))
    ActiveSheet.Sort.Header = xlYes
    ActiveSheet.Sort.Apply


' Create headings in col I thru Q
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    
    


'Begin filling Column I with stock symbols
    'get start of year price
    OpenPrice = Cells(2, 3).Value

    For i = 2 To lastrow
    
        'checking to see if stk symbol has changed to the next one
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'add the current symbol to the list
            Range("I" & SummaryTableRow).Value = Cells(i, 1).Value
            'add stock total to the list
            Range("L" & SummaryTableRow).Value = StockTotal + Cells(i, 7)
            'get end of year price
            ClosePrice = Cells(i, 6).Value
            'calc yearly change
            YearlyChange = ClosePrice - OpenPrice
            Range("J" & SummaryTableRow).Value = YearlyChange
            'calc percent change
            PercentChange = (ClosePrice - OpenPrice) / (OpenPrice + 0.0000001)
            Range("K" & SummaryTableRow).Value = PercentChange
            
            'add a line for the next symbol
            SummaryTableRow = SummaryTableRow + 1
            'reset stock total for the next symbol
            StockTotal = 0
            'reset open price for the next stl symbol
            OpenPrice = Cells(i + 1, 3).Value
        
        Else 'if symbol stays the same...
            StockTotal = StockTotal + Cells(i, 7)
            
        End If
    Next i


'color the cells
    For i = 2 To Cells(Rows.Count, 10).End(xlUp).Row
        If Cells(i, 10).Value > 0 Then Cells(i, 10).Interior.ColorIndex = 4
        If Cells(i, 10).Value < 0 Then Cells(i, 10).Interior.ColorIndex = 3
    Next i

'scan column L for greatest % incr & decr and lookup its symbol
Range("Q2").Value = Application.WorksheetFunction.Max(Range("K2", Range("K2").End(xlDown)))
Range("Q3").Value = Application.WorksheetFunction.Min(Range("K2", Range("K2").End(xlDown)))
Range("Q4").Value = Application.WorksheetFunction.Max(Range("L2", Range("L2").End(xlDown)))

'symbol for greatest % incr
pctincstk = Application.WorksheetFunction.Match(Range("Q2").Value, Range("K1", "K" & lastrow), 0)
Range("P2").Value = Cells(pctincstk, 9).Value

'symbol for greatest % decr
pctdecstk = Application.WorksheetFunction.Match(Range("Q3").Value, Range("K1", "K" & lastrow), 0)
Range("P3").Value = Cells(pctdecstk, 9).Value

'symbol for greatest stk vol
stkstkvol = Application.WorksheetFunction.Match(Range("Q4").Value, Range("L1", "L" & lastrow), 0)
Range("P4").Value = Cells(stkstkvol, 9).Value

'now make it all pretty
Range("K:K").Columns.NumberFormat = "0.00%"
Range("Q2:Q3").Columns.NumberFormat = "0.00%"
Range("A1:Q1").Font.Bold = True
Range("A:Q").Columns.AutoFit

Next a
End Sub