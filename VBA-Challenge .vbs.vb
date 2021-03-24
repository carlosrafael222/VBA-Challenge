Sub Loop_Stock_Ticker()
 
 'Name my variables


Dim Ticker as String

Dim PriceChange As Double
 PriceChange = 0

Dim YearOpen As Double
YearOpen = Cells(2, 3).Value

Dim YearClose As Double
  YearClose = 0

Dim PercentChange As Double
 PercentChange = 0
 
Dim TickerSymbol As Integer
TickerSymbol = 2

Dim TotalVolume As LongLong
 TotalVolume = 0



    ' create the table for new values
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Price_Change"
    Range("k1").Value = "Percent_Change"
    Range("l1").Value = "Total_Stock_Volume"
    
    Range("i1").Font.Bold = True
    Range("j1").Font.Bold = True
    Range("k1").Font.Bold = True
    Range("l1").Font.Bold = True

    ' create my loop
 LastRow = Cells(Rows.Count, 1).End(xlUp).Row - 1

    For i = 2 to LastRow
 
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value

            TotalVolume = TotalVolume + Cells(i, 7).Value

            YearClose = Cells(i, 6).Value

            Range("I" & TickerSymboL).Value = Cells(i, 1).Value
    'Range("J")
    'Range("K")
    'Range("L" & TotalVolume).Value = TotalVolume
    'PriceChange = YearClose - YearOpen
    'PercentChange = (PriceChange / YearOpen) * 100
    

          TickerSymbol = TickerSymbol + 1 

          TotalVolume = TotalVolume + 1

          TotalVolume = 0
          PriceChange = 0
          PercentChange = 0
   
        End If
        
    Next i 


End Sub
