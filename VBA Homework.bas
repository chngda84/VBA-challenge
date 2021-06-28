Attribute VB_Name = "Module1"
Sub Stocks_HW()
For Each ws In Worksheets
    Dim WorksheetName As String
    Dim Ticker As String
    Dim TotalStockVol As Double
    Dim OpeningStock As Double
    Dim ClosingStock As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim FirstTickerRow As Double
    Dim Summary_Table_Row As Integer
    'Create headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest % Total Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    'Set initial values
    Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    Summary_Table_Row = 2
    TotalStockVol = 0
    WorksheetName = ws.Name
    'PctIncrease
    ws.Range("O2") = ws.Cells(2, 9).Value
    ws.Range("P2") = ws.Cells(2, 11).Value
    'PctDecrease
    ws.Range("O3") = ws.Cells(2, 9).Value
    ws.Range("P3") = ws.Cells(2, 11).Value
    'TotalVol
    ws.Range("O4") = ws.Cells(2, 9).Value
    ws.Range("P4") = ws.Cells(2, 12).Value
    'Create loop
    For i = 2 To Lastrow
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            FirstTickerRow = i
        End If
        'If next cell is the same as This cell then
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Ticker symbol
            Ticker = ws.Cells(i, 1).Value
            'Total stock volume
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            'Opening stock
            OpeningStock = ws.Cells(FirstTickerRow, 3).Value
            'Closing stock
            ClosingStock = ws.Cells(i, 6).Value
            'YearlyChange
            YearlyChange = ClosingStock - OpeningStock
            'PercentChange
            'note assumptions to cater for 0 denominator
                If OpeningStock = 0 And ClosingStock = 0 Then
                    PercentChange = 0
                    ElseIf OpeningStock = 0 And ClosingStock <> 0 Then
                        PercentChange = YearlyChange
                        Else:
                            PercentChange = YearlyChange / OpeningStock
                End If
            'Output
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            ws.Range("L" & Summary_Table_Row).Value = TotalStockVol
            'Readjust Values
            Summary_Table_Row = Summary_Table_Row + 1
            TotalStockVol = 0
        Else
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
        End If
    Next i
    'Determine last row of summary table (faster to process)
    LastrowCalc = ws.Cells(Rows.Count, "J").End(xlUp).Row
    For i = 2 To LastrowCalc
        'format style
         ws.Cells(i, 11).Style = "Percent"
         ws.Cells(i, 10).Style = "Currency"
        'format colour
         If ws.Cells(i, 10).Value >= 0 And ws.Cells(i, 10).Value <> "" Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(i, 10).Interior.ColorIndex = 0
        End If
    Next i
    'Challenge
    For i = 2 To LastrowCalc
        'PctIncrease
        If ws.Cells(i + 1, 11).Value > ws.Cells(i, 11).Value And ws.Cells(i + 1, 11).Value > ws.Range("P2").Value Then
            ws.Range("O2").Value = ws.Cells(i + 1, 9).Value
            ws.Range("P2").Value = ws.Cells(i + 1, 11).Value
        End If
        'PctDecrease
        If ws.Cells(i + 1, 11).Value < ws.Cells(i, 11).Value And ws.Cells(i + 1, 11).Value < ws.Range("P3").Value Then
            ws.Range("O3").Value = ws.Cells(i + 1, 9).Value
            ws.Range("P3").Value = ws.Cells(i + 1, 11).Value
        End If
        'TotalVol
        If ws.Cells(i + 1, 12).Value > ws.Cells(i, 12).Value And ws.Cells(i + 1, 12).Value > ws.Range("P4").Value Then
            ws.Range("O4").Value = ws.Cells(i + 1, 9).Value
            ws.Range("P4").Value = ws.Cells(i + 1, 12).Value
        End If
    Next i
    'Format Cells
    ws.Range("P2").Style = "Percent"
    ws.Range("P3").Style = "Percent"
Next ws
End Sub
