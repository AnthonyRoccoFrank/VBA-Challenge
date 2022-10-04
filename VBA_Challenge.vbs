Sub VBA_Challenge()

    ' Set a variable for specifying the column of interest
  Dim Ticker As String
  
  'Set the initial variable for holding total stock volume
  Dim Stock_Total As Double
  Stock_Total = 0
  
  Dim Ticker_Column As Integer
  Ticker_Column = 2
  
  'Set New Variables to find yearly change and percent change
   Dim Open_Price As Double
   Open_Price = 0
   
   Dim Close_Price As Double
   Close_Price = 0
   
   Dim Yearly_Change As Double
   Yearly_Change = 0
   
   Dim Percent_Change As Double
   Percent_Change = 0
  
  'Create the column headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Initialize Open Price for loop below
    Open_Price = Cells(2, 3).Value

  ' Loop through rows in the column
  For i = 2 To 22771

    ' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker
      Ticker = Cells(i, 1).Value
      
      'Calculate Yearly_Change and Percent_Change
      Close_Price = Cells(i, 6).Value
      Yearly_Change = Close_Price - Open_Price
      'Check division by 0
      If Open_Price <> 0 Then
          Percent_Change = (Yearly_Change / Open_Price) * 100
      End If
      
      'Print Yearly Change and Percent Change
      Range("J" & Ticker_Column).Value = Yearly_Change
      Range("K" & Ticker_Column).Value = (CStr(Percent_Change) & "%")
      
      'Fill Yearly Change Colors
      If (Yearly_Change > 0) Then
          Range("J" & Ticker_Column).Interior.ColorIndex = 4
      ElseIf (Yearly_Change < 0) Then
          Range("J" & Ticker_Column).Interior.ColorIndex = 3
      Else
          Range("J" & Ticker_Column).Interior.ColorIndex = 2
      
      End If
      
      'Add the stock volume
      Stock_Total = Stock_Total + Cells(i, 7).Value
      
      ' Print the ticker value in column I
      Range("I" & Ticker_Column).Value = Ticker
      
      'Print the stock volume to column 11
      Range("L" & Ticker_Column).Value = Stock_Total
      
      ' Add one to the ticker column
      Ticker_Column = Ticker_Column + 1
      
      'Reset the stock total
      Stock_Total = 0
      
    'If the cell immediately following a row is the same ticker
    Else
    
        'Add to the stock total
        Stock_Total = Stock_Total + Cells(i, 7).Value

    End If

  Next i

End Sub
