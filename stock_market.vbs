Sub stock_market()

' * Create a script that will loop through all the stocks for one year for each run and take the following information.

' * The ticker symbol, Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

' * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

' * The total stock volume of the stock.

' Assumption: Records in each sheet are pre-sorted by ticker and date in ascending order


  ' Loop through all sheets
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

 ' Find the last row of each worksheet
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

  ' Create variables
  Dim ticker As String
  
  Dim yearly_change As Double
 
  Dim open_price As Double
  
  Dim close_price As Double
  
  Dim percent_change As Double
  
  Dim stock_vol As Double
  
  stock_vol = 0
 
 ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
 ' Headings
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"

 'Set Initial open price
 open_price = Cells(2, 3).Value
 
  ' Loop through each individual sheet
    
  For i = 2 To lastRow

    ' Check if we are still within the same ticker, if it is not

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker
      ticker = Cells(i, 1).Value
      
      ' Set the close price
      close_price = Cells(i, 6).Value

      ' Calculate yearly_change
       yearly_change = close_price - open_price

      ' Calculate percent change
      If (open_price = 0 And close_price = 0) Then
        percent_change = 0
        
      ElseIf (open_price = 0 And close_price <> 0) Then
        percent_change = 1
        
      Else
        percent_change = yearly_change / open_price
                
       End If
      
      
      ' Calculate volume
       stock_vol = stock_vol + Cells(i, 7).Value

      ' Print the ticker
      Range("I" & Summary_Table_Row).Value = ticker

      ' Print the yearly change
      Range("J" & Summary_Table_Row).Value = yearly_change

      ' Print the percent change
      Range("K" & Summary_Table_Row).Value = Format(percent_change, "Percent")

      ' Print the volume
      Range("L" & Summary_Table_Row).Value = stock_vol

      ' Add one to the summary table row to increment position
      
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Save the open price for new ticker
       open_price = Cells(i + 1, 3)
       
       ' Reset volume
         stock_vol = 0

    ' If the cell immediately following a row is the same ticker
    Else

      ' Add to volume
       stock_vol = stock_vol + Cells(i, 7).Value

    End If

  Next i

     ' Ticker Color coding - conditional formatting that will highlight positive change in green and negative change in red

     ' Find the last row of summary table
     
      lastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row

     ' Loop through summary table and assign colors
     
     For j = 2 To lastRowSummary

      If (Cells(j, 10).Value >= 0) Then
      
         Cells(j, 10).Interior.ColorIndex = 4
      Else
         Cells(j, 10).Interior.ColorIndex = 3
      End If

     Next j
     

 ' Stocks with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume"

 ' Set summary table fields
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

  ' Loop through the calulated summary table to find the different percentages
  
' Greatest % Increase
   For k = 2 To lastRowSummary

   If (Cells(k, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastRowSummary))) Then
    Cells(2, 16).Value = Cells(k, 9).Value
    Cells(2, 17).Value = Format(Cells(k, 11).Value, "Percent")

' Exit the loop after first match to handle multiple matches
   Exit For

        End If

    Next k

' Greatest % Decrease
  For k = 2 To lastRowSummary
   If (Cells(k, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastRowSummary))) Then
    Cells(3, 16).Value = Cells(k, 9).Value
    Cells(3, 17).Value = Format(Cells(k, 11).Value, "Percent")

' Exit the loop after first match to handle multiple matches
   Exit For

        End If

    Next k

' Greatest total volume
  For k = 2 To lastRowSummary
   If (Cells(k, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastRowSummary))) Then
    Cells(4, 16).Value = Cells(k, 9).Value
    Cells(4, 17).Value = Cells(k, 12).Value

' Exit the loop after first match to handle multiple matches
  Exit For

        End If

    Next k

Next ws


End Sub