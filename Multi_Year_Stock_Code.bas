Attribute VB_Name = "Module1"
Sub Homework_VBA()

    'Name header cells
    
    [I1].Value = "Ticker"
    [J1].Value = "Yearly Change"
    [K1].Value = "Percent Change"
    [L1].Value = "Total Stock Volume"
    


  ' Set an initial variable for holding the stock name
  Dim stock As String

  ' Set an initial variable for holding the total per stock brand
  Dim stock_volume As Double
  stock_volume = 0

  ' Set an initial variable for holding the open value per stock brand
  Dim stock_open As Double
  stock_open = [C2].Value

  ' Declare the close value per stock brand
  Dim stock_close As Double

  ' Keep track of the location for each stock brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Find last row
    'declare variable
    Dim i, lastRow As Long
        
    'calculate last row
    lastRow = Range("A" & Rows.Count).End(xlUp).Row

  ' Loop through all stock opening and closing
  For i = 2 To lastRow

    ' Check if we are still within the same stock, else:
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        
        '''Value assignment
      ' Set the Brand name
      stock = Cells(i, 1).Value

      ' Add to the Stock Total
      stock_volume = stock_volume + Cells(i, 7).Value
      
      'Set Stock close value
      stock_close = Cells(i, 6).Value
      
      
      

        '''Summary Table Fill

      ' Print the Stock Brand in the Summary Table
      Range("I" & Summary_Table_Row).Value = stock

      ' Print the Stock volume total to the Summary Table
      Range("L" & Summary_Table_Row).Value = stock_volume

      ' Print the Stock Yearly Change to the Summary Table
      Range("J" & Summary_Table_Row).Value = stock_close - stock_open
      If stock_close - stock_open >= 0 Then
             Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      Else
             Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      ' Print the Stock Percent Change to the Summary Table
      Range("K" & Summary_Table_Row).Value = (stock_close - stock_open) / stock_open
      
      
      'Set as Percent
      Range("K" & Summary_Table_Row).Value = FormatPercent(Range("K" & Summary_Table_Row))

    
    

        '''Reset
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Reset the open value
      stock_open = Cells(i + 1, 3).Value
      
      ' Reset the Brand Total
      stock_volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Stock Volume Total
      stock_volume = stock_volume + Cells(i, 7).Value

    End If

  Next i

    
End Sub

