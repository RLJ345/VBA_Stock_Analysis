Sub StockMarket()
# Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.

# You will also need to display the ticker symbol to coincide with the total volume.

# Your result should look as follows (note: all solution images are for 2015 data).
'STEP 1: Definitions
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
Dim openPrice As Double
Dim closePrice As Double
Dim tickerName As String
Dim Volume As Double

Volume = 0
Dim Row As Double
Row = 2
Dim Column As Integer
Column = 1
Dim i As Long
lastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'STEP 2: Define Headers and format cells
    Cells(1, "I").Value = "Ticker"
    Cells(1, "L").Value = "Total Stock Volume"

     
'STEP 3: Loop through Ticker Symbols
        
        For i = 2 To lastRow
            'Only count if Same Symbol Stops Counting if the Symbol Changes
            If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
                'Ticker name
                tickerName = Cells(i, "A").Value
                Cells(Row, "I").Value = tickerName
                
                
                'Total Stock Volumn
                Volume = Volume + Cells(i, "G").Value
                Cells(Row, "L").Value = Volume
               
                openPrice = Cells(i + 1, "C")
                'Counters
                Volume = 0
                Row = Row + 1
                Cells(Row, "K").NumberFormat = "0.00%"
            
            Else
                Volume = Volume + Cells(i, "G").Value
            End If
        Next i
        
    Next WS
        
End Sub

