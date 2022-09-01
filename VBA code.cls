VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StocksValue():

'loop for all worksheets
For Each ws In Worksheets

    ' Summary table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' variables to iterate through the rows
    Dim tickers As String
    Dim stockvol As Double
    Dim startrow As Double
    Dim endrow As Double
    Dim yearopen As Double
    Dim yearclose As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    
    ' tally up the stock volume
    Dim totalstockvol As Double
    
    ' write out the summary here
    Dim SummaryRow As Integer
    SummaryRow = 2
    
    'Find the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ' loop through all the rows in ticker column
    For i = 2 To LastRow
        
        ' Set conditionals to get yearopan and yearclose values for each ticker type
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            startrow = i
            yearopen = ws.Cells(startrow, 3).Value
        End If
        
        endrow = 0
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            endrow = i
            yearclose = ws.Cells(endrow, 6).Value
        End If
       
       
        ' assign values to column 1 tickers and column 7 stock volume
        tickers = ws.Cells(i, 1).Value
        stockvol = ws.Cells(i, 7).Value

        
        ' If ticker is the same then add to the volume
        totalstockvol = totalstockvol + stockvol
        
        
        ' If ticker changes, new ticker to output in summary table
        If ws.Cells(i + 1, 1).Value <> tickers Then
                
            yearlychange = (yearclose - yearopen)
            
            ws.Cells(SummaryRow, 9).Value = tickers
            ws.Cells(SummaryRow, 10).Value = yearlychange
            ws.Cells(SummaryRow, 11).Value = percentchange
            ws.Cells(SummaryRow, 12).Value = totalstockvol
        
            'Avoid divison by 0
            If (yearopen = 0) Then
                percentchange = 0
                
            Else
                percentchange = (yearlychange / yearopen) / 100
            
            End If
            
            'Change colours in Yearly Change and Percentage Change columns in summary table
            If ws.Cells(SummaryRow, 10).Value >= 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                
            Else
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            
            End If
            
            
            ' Set each row of summary table
            SummaryRow = SummaryRow + 1
            
            ' Ticker reset
            totalstockvol = 0
            yearlychange = 0
            
    
        End If
  
    
    Next i
    
    
    ' Number formatting
    ws.Range("J:J").NumberFormat = "0.00"
    ws.Range("K:K").NumberFormat = "0.00%"
    
    'Create secondary summary table

        'Define ranges from primary summary table (to be used in secondary table)
            
        sum_TickerCol = ws.Range("I:I")
        sum_YPercentChangeCol = ws.Range("K:K")
        sum_TotalVolCol = ws.Range("L:L")
    
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
    
        'Find grt increase nad grt decrease values
        greatestPerInc = Application.WorksheetFunction.Max(sum_YPercentChangeCol)
        greatestPerDec = Application.WorksheetFunction.Min(sum_YPercentChangeCol)
        greatedTotalVol = Application.WorksheetFunction.Max(sum_TotalVolCol)

        'Find corresponding ticker symbol using index function
        greatestPercenIncrease = WorksheetFunction.Index(sum_TickerCol, WorksheetFunction.Match(greatestPerInc, sum_YPercentChangeCol, 0))
        greatestPercenDecrese = WorksheetFunction.Index(sum_TickerCol, WorksheetFunction.Match(greatestPerDec, sum_YPercentChangeCol, 0))
        greatedTotalVolume = WorksheetFunction.Index(sum_TickerCol, WorksheetFunction.Match(greatedTotalVol, sum_TotalVolCol, 0))

        'Print outputs into secondary summary table
        ws.Cells(2, 15).Value = greatestPercenIncrease
        ws.Cells(2, 16).Value = greatestPerInc

        ws.Cells(3, 15).Value = greatestPercenDecrese
        ws.Cells(3, 16).Value = greatestPerDec

        ws.Cells(4, 15).Value = greatedTotalVolume
        ws.Cells(4, 16).Value = greatedTotalVol
        
        ws.Cells(2, 14).NumberFormat = "0.00" + "%"
        ws.Cells(3, 14).NumberFormat = "0.00" + "%"
        ws.Cells(4, 14).NumberFormat = "###,###,###,#00"


Next ws

End Sub

    
    
    
