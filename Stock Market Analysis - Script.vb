Sub stock_market_analysis():
 
	Dim complete_ws As Worksheet
	  For Each complete_ws In ThisWorkbook.Worksheets
           complete_ws.Activate

            Cells.EntireColumn.AutoFit

   	 Dim stocks As Integer
  	 Dim total_volume As Double
 	 Dim i As Long
   	 Dim yearly_change As Double
  	 Dim percent_change As Double

       		  stocks = 2
       		  total_volume = 0
       		  yearly_change = 0
       		  percent_change = 0
        
    Dim opening_price As Double
    Dim closing_price As Double

   'get the value of first opening price
   opening_price = Cells(2, 3).Value
   
   'set the summary table headers
  	 Cells(1, 9).Value = "Ticker"
   	 Cells(1, 10).Value = " Yearly Change"
  	 Cells(1, 11).Value = "Percent Change"
   	 Cells(1, 12).Value = "Total Stock Volume"

   'go through every row
  For i = 2 To Range("A2").End(xlDown).Row
  
   total_volume = Cells(i, 7).Value + total_volume
    'when ticker changes to next stock

     If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        'add to summary table and reset volume
         Cells(stocks, 12).Value = total_volume
         total_volume = 0
        
        'add current stock name to summary table
         Cells(stocks, 9).Value = Cells(i, 1).Value
        
        'get the value of current closing price
         closing_price = Cells(i, 6).Value
        
        'calculate yearly change and add to summary table
         yearly_change = closing_price - opening_price
         Cells(stocks, 10).Value = yearly_change
         
        'calculate percentage change and add to summary table
         percent_change = yearly_change / opening_price
         Cells(stocks, 11).Value = percent_change
        'get the value of new opening price
         opening_price = Cells(i + 1, 3).Value
         
        'move to the next stock on summary table
         stocks = stocks + 1
        
       End If
     
    Next i
     
      	For i = 2 To Range("J2").End(xlDown).Row
     	  If Cells(i, 10) < 0 Then

      		 Cells(i, 10).Interior.ColorIndex = 3
      		 Cells(i, 11).Interior.ColorIndex = 3

      	  ElseIf Cells(i, 10) >= 0 Then

      		 Cells(i, 10).Interior.ColorIndex = 4
   		 Cells(i, 11).Interior.ColorIndex = 4
  	 End If
   Next i
           
      For i = 2 To Range("K2").End(xlDown).Row
      
       Cells(i, 11).NumberFormat = "0.00%"
        
   Next i
       
    	 Dim greatest_increase_percent As Double
   	 Dim greatest_decrease_percent As Double
   	 Dim greatest_volume As Double
    
    	   Cells(2, 14).Value = "Greatest % Increase"
    	   Cells(3, 14).Value = "Greatest % Decrease"
   	   Cells(4, 14).Value = "Greatest Total Volume"
   	   Cells(1, 15).Value = "Ticker"
   	   Cells(1, 16).Value = "Value"

     x = Range("K2").EntireColumn

   	 greatest_increase_percent = Application.WorksheetFunction.Max(x)
   		 Cells(2, 16).Value = greatest_increase_percent
  		  Cells(2, 16).NumberFormat = "0.00%"
   	 greatest_decrease_percent = Application.WorksheetFunction.Min(x)
   		 Cells(3, 16).Value = greatest_decrease_percent
   		  Cells(3, 16).NumberFormat = "0.00%"

    Z = Range("L2").EntireColumn
     
    	greatest_volume = Application.WorksheetFunction.Max(Z)
  		  Cells(4, 16).Value = greatest_volume
 
   	 increase = Range("P2").Value
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    	 If Cells(i, 11).Value = increase Then
    	  found_ticker = Cells(i, 9).Value
    
      End If
     
    Next i
     
       Cells(2, 15).Value = found_ticker
     
    	increase = Range("P3").Value
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 11).Value = increase Then
          found_ticker = Cells(i, 9).Value
    
     End If
     
     Next i
     
     Cells(3, 15).Value = found_ticker
    
    increase = Range("P4").Value
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
     If Cells(i, 12).Value = increase Then
     found_ticker = Cells(i, 9).Value
    
 	 End If
     
  Next i
     
     Cells(4, 15).Value = found_ticker
     
   Cells.EntireColumn.AutoFit

Next complete_ws
    
End Sub