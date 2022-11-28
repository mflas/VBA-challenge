Attribute VB_Name = "Module1"
Sub MultiYear()
Attribute MultiYear.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MultiYear Macro
'
'
 '-------------------------------
 ' LOOP THROUGH ALL WORKSHEETS


Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

 '-------------------------------------------
 
    
    'Set an initial variable for holding the ticker
    Dim Ticker_Name As String
    
    'Set a variable to hold the opening price
    Dim Opening_Price As Currency
    
     'Set a variable to hold the closing price
    Dim Closing_Price As Currency
    
     'Set a variable to hold the Yearly change
    Dim Yearly_Change As Currency
    

' Set an initial variable for hodling the total volume
    Dim Ticker_Volume As LongLong
    
'Set volume to zero for  start
    Ticker_Volume = 0
    
'   Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
'Set last row for loop using last row
    Dim LastRow As Long
    LastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    
'Print the table headers and first opening price
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Opening Price"
    ws.Range("K1").Value = "Closing Price"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percentage Change"
    ws.Range("N1").Value = "Total Stock Volume"
    ws.Range("J2").Value = ws.Range("C2").Value
    
        
'Loop through all the tickers

    For I = 2 To LastRow

    ' Check if we still in the same ticker
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
    
    'set the ticker name
    Ticker_Name = ws.Cells(I, 1).Value
    
    'set the closing price
    Closing_Price = ws.Cells(I, 6).Value
 
  'Set the opening price
    Opening_Price = ws.Cells(I + 1, 3).Value
 
 'Print the Opening price in the summary table
    ws.Range("J" & Summary_Table_Row + 1).Value = Opening_Price
 

    ' Add to the ticker volume
    Ticker_Volume = Ticker_Volume + ws.Cells(I, 7).Value
    
    'Print the ticker name in the summary table
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    
       
    'Print the Close price in the summary table
    ws.Range("K" & Summary_Table_Row).Value = Closing_Price

 'Add the Yearly change into the table for Yearly change and Percentage change
    ws.Range("L" & Summary_Table_Row).Value = (Closing_Price - ws.Range("J" & Summary_Table_Row).Value)
    
 
 'Add the Percentage change into the table
    ws.Range("M" & Summary_Table_Row).Value = ws.Range("L" & Summary_Table_Row).Value / ws.Range("J" & Summary_Table_Row).Value
    ws.Range("M" & Summary_Table_Row).NumberFormat = "0.00%"
      
    'Print the ticker volume in the summary table
    ws.Range("N" & Summary_Table_Row).Value = Ticker_Volume
    
    ' Add 1 to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    'Reset the Ticker volume
    Ticker_Volume = 0
    
      
    ' If the cell immediately following a row is the same ticker
    Else
    
    ' Add to the Ticker Volume
    Ticker_Volume = Ticker_Volume + ws.Cells(I, 7).Value
    
 
     
    End If
    
  Next I
  
'This section applies the conditional formatting
'Define Range

    Dim MyRange As Range
    Set MyRange = ws.Range("L2:L5000")

    'Delete Existing Conditional Formatting from Range
    MyRange.FormatConditions.Delete

  'Defining and setting the criteria for each conditional format
   Set condition1 = MyRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
   Set condition2 = MyRange.FormatConditions.Add(xlCellValue, xlLess, "=0")

   'Defining and setting the format to be applied for each condition
   With condition1
    .Interior.Color = vbGreen
    .Font.Bold = True
   End With

   With condition2
     .Interior.Color = vbRed
      .Font.Bold = True
   End With

'Delete the price columns
ws.Columns(10).EntireColumn.Delete
ws.Columns(10).EntireColumn.Delete


Next ws

  
End Sub



