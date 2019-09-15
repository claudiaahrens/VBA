Sub StockTickerMacro()
    ' Set initial variable for stock ticker
    Dim Stock_Ticker_Symbol As String
    
    ' Set an initial variable for holding the total per stock ticker
    Dim Stock_Ticker_Total As Double
    Stock_Ticker_Total = 0
    
    ' Set to determine which row to write out ticker information
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
        
    For i = 2 To 760193
      ' Check if we are still within the same stock ticker type, if it is not...
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
      ' The next row Stock ticker is different. Need to include current rows stock volume
       Stock_Ticker_Total = Stock_Ticker_Total + Cells(i, 7).Value
       
       ' What is the current Ticker Symbol Were Looking At
       Stock_Ticker_Symbol = Cells(i, 1).Value
       
       ' Need to account for current_rows Total Volume
       Range("I" & Summary_Table_Row).Value = Stock_Ticker_Symbol
       Range("J" & Summary_Table_Row).Value = Stock_Ticker_Total
       
       ' Update the Summary Row for the next run
       Summary_Table_Row = Summary_Table_Row + 1
       
       ' Reset Stock_Ticker_Total for the next run
       Stock_Ticker_Total = 0
      Else
        ' We are looking at the same Stock Ticker Symbol
        Stock_Ticker_Total = Stock_Ticker_Total + Cells(i, 7).Value
      End If
      
    Next i
    

End Sub