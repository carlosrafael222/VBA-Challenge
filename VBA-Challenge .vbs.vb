'Create Subroutine
Sub StockVolume()


'Assign variables

Dim ws As Worksheet
Dim Ticker As String
Dim TotalVolume As LongLong
Dim tickersymbol As Integer
Dim PercentChange As Double
Dim YearlyChange As Double
Dim YearlyOpen As Double
Dim YearlyClose As Double

'Create new Table on all worksheets
For Each ws In Worksheets
    TotalVolume = 0
    tickersymbol = 0
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1") = "YearlyChange"
    ws.Range("K1") = "PercentChange"
    ws.Range("L1").Value = "Total Stock Volume"
    
    Range("i1").Font.Bold = True
    Range("j1").Font.Bold = True
    Range("k1").Font.Bold = True
    Range("l1").Font.Bold = True

'count to the last row of the columns
lastrow = Cells(Rows.Count, "A").End(xlUp).Row - 1

'Create loop
  For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'set ticker
        Ticker = ws.Cells(i, 1).Value
        
'set stock total volume
         TotalVolume = TotalVolume + ws.Cells(i, 7).Value
'set Yearly Close
         YearlyClose = ws.Cells(i, 6).Value
'set Yearly Open
         YearlyOpen = ws.Cells(i, 3).Value
'set Yearly Change
         YearlyChange = (YearlyClose - YearlyOpen)
'Add the percent change
                If YearlyOpen = 0 Then
            
                Else
                PercentChange = (YearlyClose - YearlyOpen) / YearlyOpen
            
                End If
'fill in the data on the new table
        ws.Range("I" & 2 + tickersymbol).Value = Ticker
        ws.Range("J" & 2 + tickersymbol).Value = YearlyChange
        ws.Range("K" & 2 + tickersymbol).Value = PercentChange
            ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Range("L" & 2 + tickersymbol).Value = TotalVolume
'reset Totalvolume
             TotalVolume = 0
'Iterate plus 1 down, reset the total for tickersymbol
             tickersymbol = tickersymbol + 1
             
       'format the cells on new table
        If ws.Range("J" & 2 + tickersymbol).Value >= 0 Then
        
            ws.Range("J" & 2 + tickersymbol).Interior.ColorIndex = 4
            
        ElseIf ws.Range("J" & 2 + tickersymbol).Value < 0 Then
        
            ws.Range("J" & 2 + tickersymbol).Interior.ColorIndex = 3
            
        End If
             
    Else
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
    End If


Next i
        'reset my count
        TotalVolume = 0
        tickersymbol = 0

Next ws
               
End Sub