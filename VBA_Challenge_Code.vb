Sub stock():
   'set table headers
   Range("I1") = "Ticker"
   Range("J1") = "Yearly Change"
   Range("K1") = "Percent Change"
   Range("L1") = "Total Stock Volume"
    
    'set variable for the ticker
    Dim ticker As String

    'set varable for open, close, and yearly change, and percent change and set to zero
    Dim stockopen As Double
    Dim stockclose As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    
    stockopen = 0
    stockclose = 0
    yearlychange = 0
    percentchange = 0
    
    
    'set variable for volume
    Dim Volume As Double
   
    'set the first open value, as if not listed there is no open price and value is end price
    stockopen = Cells(2, 3).Value
    
    'keep track of the location of the summary table
    Dim Summary_Table As Integer
    Summary_Table = 2
    
    'create a count for the rows at the end of the worksheet
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
  
     'loop through all rows in the worksheet
     For i = 2 To last_row
     
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then ' found the boundary
     
     'set the ticker name
     ticker = Cells(i, 1).Value
     
   
     'calculate the yearly change
     stockclose = Cells(i, 6).Value
     yearlychange = stockclose - stockopen
     
     'calculate the percent change
        'add if statement to check for dividing by zero
       If stockopen <> 0 Then
        percentchange = (yearlychange / stockopen) * 100
        End If
        
    
     'add volume together
     Volume = Volume + Cells(i, 7).Value
     
     'write the ticker in the table
     Range("I" & Summary_Table).Value = ticker
     
     'write the yearly change in the table
     Range("J" & Summary_Table).Value = yearlychange
     
     'add colors for positve and negative percent changes
     
         'fill green if positive
        If (yearlychange > 0) Then
         Range("J" & Summary_Table).Interior.ColorIndex = 4
     
        'fill red if negative change
         ElseIf (yearlychange <= 0) Then
        Range("j" & Summary_Table).Interior.ColorIndex = 3
     
         End If
     
     'write the percent change to the table and format the percent into percentage vs a number with lots of decimals
     Range("K" & Summary_Table).Value = (CStr(percentchange) & "%")
     
     'write the volume to the table
     Range("L" & Summary_Table).Value = Volume
      
     'add one to the summary table to print next line
     Summary_Table = Summary_Table + 1
     
     'reset the values back to zero once complete
      percentchange = 0
      yearlychange = 0
     Volume = 0
     
     'create new open price for stocks
     stockopen = Cells(i + 1, 3).Value
     
     Else
     'add volumes until ticker symbol changes
     Volume = Volume + Cells(i, 7).Value
         
     End If
     
     Next i
     
End Sub
