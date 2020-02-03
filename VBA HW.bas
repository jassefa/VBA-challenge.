Attribute VB_Name = "Module1"
Sub HW_VBA_2()

Dim ws As Worksheet

For Each ws In Sheets

'Set an initial variable
Dim Percent_change As Variant
Dim Yearly_change As Double
Dim Total_stock_Volume As Double
Dim ticker As String
Dim Summary_table_row As Double
    Summary_table_row = 2
Dim open_date As Double
    open_date = Cells(2, 3)

' Add new Column Header
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
        
    With ActiveSheet
    lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
        
    ActiveSheet.Columns("K").NumberFormat = "0.00%"
        
Dim i As Double
        
    'Loop through all
    For i = 2 To lastrow
    Dim close_date As Double
      
    'Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
     
     'Set Ticker name
    ticker = Cells(i, 1).Value
     
     'Get yearly change caluclated
    close_date = Cells(i, 6).Value
     
    Yearly_change = close_date - open_date
     
        If open_date = 0 Then
        Percent_change = 0
     
     'Print the Percent Change to the Summary Table
    Range("K" & Summary_table_row).Value = Percent_change
     
      Else
    
    Percent_change = Yearly_change / open_date
    
    Range("K" & Summary_table_row).Value = Percent_change
      
 End If
     
     open_date = Cells(i + 1, 3).Value
    
    'add to the total stock volume
     Total_stock_Volume = Total_stock_Volume + Cells(i, 7).Value
     
    ' Print the Ticker name in the Summary Table
    Range("I" & Summary_table_row).Value = ticker
    
    'Print the Total Stock Volume to the Summary Table
    Range("L" & Summary_table_row).Value = Total_stock_Volume
    
    'Print the Yearly Change to the Summary Table
    Range("J" & Summary_table_row).Value = Yearly_change
    
    'add one to the summary table row
    Summary_table_row = Summary_table_row + 1
    
    'reset the total stock volume and yearly change
   Total_stock_Volume = 0
   
'If the cell immediately following a row is the same ticker...

    Else
'add to the Total stock volume
Total_stock_Volume = Total_stock_Volume + Cells(i, 7).Value

    End If
 
 If Cells(i, 10) >= 0 Then
 Cells(i, 10).Interior.ColorIndex = 4
 
 Else
 Cells(i, 10).Interior.ColorIndex = 3
 End If
 
Next i

ws.Activate

Next

End Sub






