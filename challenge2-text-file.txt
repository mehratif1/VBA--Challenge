
Sub Stocks()

Dim ticker As String
Dim closing_price As Double
Dim opening_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Roundnumber As Double
Dim maxvalue As Double
Dim minvalue As Double
Dim r1 As Range
Dim r2 As Range

Dim maxvolume As Double

'loop through all sheets

For Each WS In Worksheets
    'Determine the Last Row
    
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
'whenever the data in row i+1 is not equal to the data in i it should print the value in cell i in the column J
'the for loop going all the way from row 2 to the last row to check the ticker names

j = 2
test = 0

'opening price of AAB
opening_price = WS.Cells(2, 3).Value
For i = 2 To LastRow
    If WS.Cells(i, 1) <> WS.Cells(i + 1, 1) Then
    
      ' Add to the total volume
      
        totalvolume = totalvolume + WS.Cells(i, 7)
        WS.Cells(j, 12).Value = totalvolume
        ticker = WS.Cells(i, 1).Value
        WS.Cells(j, 9).Value = ticker
        
        closing_price = WS.Cells(i, 6).Value
        yearly_change = closing_price - opening_price
        WS.Cells(j, 10).Value = yearly_change
        
'calculating percent change

        percent_change = (yearly_change / opening_price) * 100
        WS.Cells(j, 11).Value = percent_change
        
        
        'rounding to 2 decimal place
        Roundnumber = Round(percent_change, 2)
        WS.Cells(j, 11).Value = "%" & Roundnumber



        Select Case yearly_change

            Case Is > 0
            WS.Range("J" & 2 + test).Interior.ColorIndex = 4
    
            Case Is < 0
            WS.Range("J" & 2 + test).Interior.ColorIndex = 3
            
            
            End Select
            
        
        Select Case percent_change
        
        
        
            Case Is > 0
            WS.Range("K" & 2 + test).Interior.ColorIndex = 4
    
            Case Is < 0
            WS.Range("K" & 2 + test).Interior.ColorIndex = 3
            
            End Select
            
        
            
        test = test + 1
        
        
    WS.Range("O3").Value = "Greatest % Increase"
    WS.Range("O4").Value = "Greatest % Decrease"
    WS.Range("O5").Value = "Greatest Total Volume"
    WS.Range("P2").Value = "Ticker"
    WS.Range("Q2").Value = "Value"
    
    'Finding the greatest% increase and the greatest % decrease
    
        Set r1 = WS.Range("K2:K" & Rows.Count)
       
        Dim minvaluestr As String
        Dim maxvaluestr As String
        
        minvalue = Application.WorksheetFunction.min(r1)
        
        minvaluestr = (minvalue * 100) & "%"
        WS.Range("Q4").Value = minvaluestr
        
        maxvalue = Application.WorksheetFunction.Max(r1)
        maxvaluestr = (maxvalue * 100) & "%"
        WS.Range("Q3").Value = maxvaluestr
  
        
    
    
    'Finding the greatest total volume
        
    
        Set r2 = WS.Range("L2:L" & Rows.Count)
        maxvolume = Application.WorksheetFunction.Max(r2)
        WS.Range("Q5").Value = maxvolume
        
        Dim maxRow As Long, minRow As Long
        maxRow = Application.WorksheetFunction.Match(maxvalue, r1, 0)
        minRow = Application.WorksheetFunction.Match(minvalue, r1, 0)
        
     ' Write the tickers corresponding to max and min percent change to the P column
        WS.Range("P3").Value = WS.Cells(maxRow + 1, 9).Value
        WS.Range("P4").Value = WS.Cells(minRow + 1, 9).Value
        
        Dim maxvol As Long
        maxvol = Application.WorksheetFunction.Match(maxvolume, r2, 0)
        WS.Range("P5").Value = WS.Cells(maxvol + 1, 9).Value
        
        
         

    
    
    

j = j + 1

opening_price = WS.Cells(i + 1, 3).Value

totalvolume = 0
'if the data in next cell is same

 Else


totalvolume = totalvolume + WS.Cells(i, 7).Value

 End If
 
 Next i

'Add the word Ticker to the new column header
WS.Cells(1, 9).Value = "Ticker Symbol"
'Add the header Yearly change
WS.Cells(1, 10).Value = "Yearly Change"
'Add the header Percent Change
WS.Cells(1, 11).Value = "Percent Change"
'Add the header Total Stock Volume
WS.Cells(1, 12).Value = "Total Stock Volume"

Next WS

End Sub










