Attribute VB_Name = "Module1"
Sub Stocks():
    For Each Sheet In Worksheets
    
        Dim stockTicker As String
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim highestPrice As Double
        Dim lowestPrice As Double
        Dim totalVolume As Double
        Dim rowIndex As Long
        Dim tickerNumber As Long

        openingPrice = Sheet.Range("C2").Value
        stockTicker = Sheet.Range("A2").Value
        highestPrice = Sheet.Range("D2").Value
        lowestPrice = Sheet.Range("E2").Value
        totalVolume = 0

        tickerNumber = 2
        Dim lastRow As Long
    
        Sheet.Range("I1").Value = "Stock Ticker"
        Sheet.Range("J1").Value = "Yearly Change"
        Sheet.Range("K1").Value = "Percent Change"
        Sheet.Range("L1").Value = "Total Stock Volume"

        lastRow = Sheet.Range("A" & Rows.Count).End(xlUp).Row
    
        For rowIndex = 2 To lastRow
        
    
            If Sheet.Range("A" & rowIndex).Value = stockTicker Then
              
                closingPrice = Sheet.Range("F" & rowIndex).Value
            
        
                totalVolume = totalVolume + Sheet.Range("G" & rowIndex).Value
     
                If highestPrice < Sheet.Range("D" & rowIndex).Value Then
                    highestPrice = Sheet.Range("D" & rowIndex).Value
                End If
        
                If lowestPrice > Sheet.Range("E" & rowIndex).Value Then
                    lowestPrice = Sheet.Range("E" & rowIndex).Value
                End If
            
            Else

                Sheet.Range("I" & tickerNumber).Value = stockTicker
        
                Sheet.Range("J" & tickerNumber).Value = closingPrice - openingPrice
            
                If (closingPrice - openingPrice) < 0 Then
                    Sheet.Range("J" & tickerNumber).Interior.ColorIndex = 3
                Else
                    Sheet.Range("J" & tickerNumber).Interior.ColorIndex = 4
                End If
            
                Sheet.Range("K" & tickerNumber).Value = (closingPrice - openingPrice) / openingPrice
            
                Sheet.Range("K" & tickerNumber).NumberFormat = "0.00%"
        
                Sheet.Range("L" & tickerNumber).Value = totalVolume
            
    
                openingPrice = Sheet.Range("C" & rowIndex).Value
                stockTicker = Sheet.Range("A" & rowIndex).Value
                highestPrice = Sheet.Range("D" & rowIndex).Value
                lowestPrice = Sheet.Range("E" & rowIndex).Value
                totalVolume = Sheet.Range("G" & rowIndex).Value
                tickerNumber = tickerNumber + 1
            End If
               
        Next rowIndex

    
        Dim maxPercentIncrease As Double
        Dim maxPercentIncreaseStock As String
        Dim maxPercentDecrease As Double
        Dim maxPercentDecreaseStock As String
        Dim maxTotalVol As Double
        Dim maxTotalVolStock As String
    

        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVol = 0
    

        lastRow = Sheet.Range("I" & Rows.Count).End(xlUp).Row
    
        For rowIndex = 2 To lastRow
            If Sheet.Range("K" & rowIndex).Value > maxPercentIncrease Then
                maxPercentIncrease = Sheet.Range("K" & rowIndex).Value
                maxPercentIncreaseStock = Sheet.Range("I" & rowIndex).Value
            End If
        
            If Sheet.Range("K" & rowIndex).Value < maxPercentDecrease Then
                maxPercentDecrease = Sheet.Range("K" & rowIndex).Value
                maxPercentDecreaseStock = Sheet.Range("I" & rowIndex).Value
            End If
        
            If Sheet.Range("L" & rowIndex).Value > maxTotalVol Then
                maxTotalVol = Sheet.Range("L" & rowIndex).Value
                maxTotalVolStock = Sheet.Range("I" & rowIndex).Value
            End If
        
        Next rowIndex
    
        Sheet.Range("O2").Value = "Greatest Percentage Increase"
        Sheet.Range("O3").Value = "Greatest Percentage Decrease"
        Sheet.Range("O4").Value = "Greatest Total Volume"
        Sheet.Range("P1").Value = "Stock Ticker"
        Sheet.Range("Q1").Value = "Value"
    
        Sheet.Range("P2").Value = maxPercentIncreaseStock
        Sheet.Range("Q2").Value = maxPercentIncrease
        Sheet.Range("Q2").NumberFormat = "0.00%"
        Sheet.Range("P3").Value = maxPercentDecreaseStock
        Sheet.Range("Q3").Value = maxPercentDecrease
        Sheet.Range("Q3").NumberFormat = "0.00%"
        Sheet.Range("P4").Value = maxTotalVolStock
        Sheet.Range("Q4").Value = maxTotalVol
    
    Next Sheet
End Sub

