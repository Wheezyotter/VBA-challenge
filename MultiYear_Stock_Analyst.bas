Attribute VB_Name = "Module11"
Sub AnalyzeStocks()
' Macro to analyze stock data and produce a summary table of each stock type
    For Each ws In Worksheets
    
        ' Declares all variables used in the macro
        Dim TotalStock As Long
        Dim Ticker As String
        Dim NextTicker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim CurrentVol As LongLong
        Dim TotalVol As LongLong
        Dim TableRow As Integer
        Dim i As Long
        Dim GreatestPerIn As Double
        Dim GeatestetPerDe As Double
        Dim GreatestTotVol As LongLong
    
        ' Calculates total number of rows per sheet
        TotalStock = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Sets variables to start the loop and greatest variables
        OpenPrice = ws.Cells(2, 3).Value
        TotalVol = 0
        TableRow = 2
        GreatestPerIn = 0
        GreatestPerDe = 0
        GreatestTotVol = 0
        
        ' Formats headers for summary table and greatest table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
       ' Loops through each stock record
        For i = 2 To TotalStock
            'Collects ticker name and stock volume for current iteration
            Ticker = ws.Cells(i, 1).Value
            NextTicker = ws.Cells(i + 1, 1).Value
            CurrentVol = ws.Cells(i, 7).Value
            
            'Calculates total volume of stock type
            TotalVol = TotalVol + CurrentVol
            
            'Determines the end of one stock type and the start of a new stock type
            If (Ticker <> NextTicker) Then
                ClosePrice = ws.Cells(i, 6).Value
                
                ' Calculates yearly and percent change for entire stock type
                YearlyChange = ClosePrice - OpenPrice
                PercentChange = (YearlyChange / OpenPrice) * 100
                
                ' Fills in stock info into the summary table
                ws.Cells(TableRow, 9).Value = Ticker
                ws.Cells(TableRow, 10).Value = YearlyChange
                ws.Cells(TableRow, 11).Value = Str(Round(PercentChange, 2)) + "%"
                ws.Cells(TableRow, 12).Value = TotalVol
                
                ' Formats percent change cells green for increasel, red for decrease
                If (YearlyChange > 0) Then
                
                    ws.Cells(TableRow, 10).Interior.ColorIndex = 4
                
                ElseIf (YearlyChange < 0) Then
    
                    ws.Cells(TableRow, 10).Interior.ColorIndex = 3
                
                End If
                
                ' Checks which percentage has the largest increase/decrease and
                ' total volume and then fills it into a summary table
                If (PercentChange > GreatestPerIn) Then
                
                    GreatestPerIn = PercentChange
                    ws.Cells(2, 16).Value = Ticker
                    ws.Cells(2, 17).Value = Str(Round(GreatestPerIn, 2)) + "%"
                    
                ElseIf (PercentChange < GreatestPerDe) Then
                
                    GreatestPerDe = PercentChange
                    ws.Cells(3, 16).Value = Ticker
                    ws.Cells(3, 17).Value = Str(Round(GreatestPerDe, 2)) + "%"
                    
                End If
                    
                If (TotalVol > GreatestTotVol) Then
                
                    GreatestTotVol = TotalVol
                    ws.Cells(4, 16).Value = Ticker
                    ws.Cells(4, 17).Value = GreatestTotVol
                    
                End If
                
                'Resets loop variables for next stock type and sets next line in summary table
                TotalVol = 0
                TableRow = TableRow + 1
                YearStart = ws.Cells(i + 1, 2).Value
                OpenPrice = ws.Cells(i + 1, 3).Value
    
            End If
        Next i
    Next
End Sub
