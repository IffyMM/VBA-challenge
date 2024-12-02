Attribute VB_Name = "Module1"
Sub MultipleYear_stock()


'create variables

Dim i As Long
Dim ticker As String, currentticker As String

Dim firstOpen As Double
Dim lastClose As Double
Dim nextindex As Long

Dim ws As Worksheet
Dim lastrow As Long

Dim QuarterlyChange As Double
Dim Percentchange As Double
Dim TotalStockVolume As Double
Dim Summary_ticker_row As Long
Dim highestincrease As Double, lowestdecrease As Double
Dim largestvolume As Double

Dim highestincrease_ticker As String
Dim lowestdecrease_ticker As String
Dim largestvolume_ticker As String



'loops through all worksheets
For Each ws In Worksheets
'declare last row
ws.Activate

'Set ws = ThisWorkbook.Sheets("Q1")
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'header columns
ws.Range("I1").Value = "ticker"
ws.Range("J1").Value = "Quarterly change"
ws.Range("K1").Value = "Percent change"
ws.Range("L1").Value = "TotalStockVolume"


 ' Remove any existing conditional formatting in Column K (Percent Change)
        ws.Range("K2:K" & lastrow).FormatConditions.Delete

'initialize variables
ticker = ""
firstOpen = 0
lastClose = 0
nextindex = 2
TotalStockVolume = 0
Summary_ticker_row = 2

    ' Loop through the rows to find the quarterly change for each stock ticker
For i = 2 To lastrow

    currentticker = ws.Cells(i, 1).Value

    
    If currentticker <> ticker Then
    If ticker <> "" Then
    
QuarterlyChange = lastClose - firstOpen

'conditional formatting quarter change with colors
If QuarterlyChange > 0 Then
ws.Cells(nextindex, 10).Interior.ColorIndex = 4
ElseIf QuarterlyChange < 0 Then
ws.Cells(nextindex, 10).Interior.ColorIndex = 3
End If
 

If firstOpen <> 0 Then
Percentchange = (QuarterlyChange / firstOpen) * 100
   Else
Percentagechange = 0
End If

ws.Cells(nextindex, 9).Value = ticker
ws.Cells(nextindex, 10).Value = QuarterlyChange
ws.Cells(nextindex, 11).Value = Percentchange / 100
ws.Cells(nextindex, 11).NumberFormat = "0.00%"
            
ws.Cells(nextindex, 12).Value = TotalStockVolume

'increment to next ticker

 nextindex = nextindex + 1
 
  End If
  
  'reset for new ticker
  ticker = currentticker
  firstOpen = ws.Cells(i, 3).Value
  lastClose = ws.Cells(i, 6).Value
  TotalStockVolume = 0
  End If
  
  
  TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
  
  lastClose = ws.Cells(i, 6).Value
  
  If i = lastrow Or ws.Cells(i + 1, 1).Value <> currentticker Then
            ' Calculate for the last ticker
            QuarterlyChange = lastClose - firstOpen
            If firstOpen <> 0 Then
                Percentchange = (QuarterlyChange / firstOpen) * 100
            Else
                Percentchange = 0 'Prevent division by zero
            End If

            ' Output the results for the last ticker
            ws.Cells(nextindex, 9).Value = ticker ' Column I: Ticker
            ws.Cells(nextindex, 10).Value = QuarterlyChange ' Column J: Quarterly Change
            ws.Cells(nextindex, 11).Value = Percentagechange ' Column K: Percentage Change
            ws.Cells(nextindex, 11).NumberFormat = "0.00%"
            
            ws.Cells(nextindex, 12).Value = TotalStockVolume ' Column L: Total Volume
            
            ws.Cells(Summary_ticker_row, 9).Value = ticker
            ws.Cells(Summary_ticker_row, 12).Value = TotalStockVolume
            ws.Cells(Summary_ticker_row, 10).Value = QuarterlyChange
            ws.Cells(Summary_ticker_row, 11).Value = Percentchange
            
            Summary_ticker_row = Summary_ticker_row + 1
             
        End If
    Next i
    
  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greatest % Decrease"
  Range("O4").Value = "Greatest Total volume"
  Range("P1").Value = "Ticker"
   Range("Q1").Value = "Value"
   
   For i = 2 To lastrow
   
   If Cells(i, 11).Value > highestincrease Then
   highestincrease = Cells(i, 11).Value
   highestincrease_ticker = Cells(i, 9).Value
   End If
   
   If Cells(i, 11).Value < lowestdecrease Then
   lowestdecrease = Cells(i, 11).Value
   lowestdecrease_ticker = Cells(i, 9).Value
   End If
   
   If Cells(i, 12).Value > largestvolume Then
   largestvolume = Cells(i, 12).Value
   largestvolume_ticker = Cells(i, 9).Value
   End If
   
   
Next i

 Range("P2").Value = Format(highestincrease_ticker, "percent")
 Range("Q2").Value = Format(highestincrease, "percent")
 
 Range("P3").Value = Format(lowestdecrease_ticker, "percent")
 Range("Q3").Value = Format(lowestdecrease, "percent")
 
 Range("P4").Value = largestvolume_ticker
 Range("Q4").Value = largestvolume
 
    MsgBox "done"
    
    Next ws
End Sub

