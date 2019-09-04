Attribute VB_Name = "Module1"
Sub stock_data_analysis_2()

    ' Make sure code works for all sheets
    
    For Each ws In Worksheets
    
    ws.Activate
  
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
   ' Define variables
   
    Dim stock_ticker As String
    Dim total_volume As Double
    total_volume = 0
    Dim summary_table_row As Integer
    summary_table_row = 2
    Dim column As Integer
    column = 1
    
    Dim i As Long
    
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    
    ' Define Headers
    
    Cells(1, "I").Value = "Stock Ticker"
    Cells(1, "J").Value = "Total Volume"
    Cells(1, "K").Value = "Yearly Change"
    Cells(1, "L").Value = "Percent Change"
    
    open_price = Cells(2, column + 2).Value
    
    For i = 2 To LastRow
    
    ' Loop through all stocks
    
    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
    
    stock_ticker = Cells(i, column).Value
    Cells(summary_table_row, column + 8).Value = stock_ticker
    
    ' Determine yearly change
    
    close_price = Cells(i, column + 5).Value
    yearly_change = close_price - open_price
    Cells(summary_table_row, column + 10).Value = yearly_change
    
    ' Determine Percentage change
    
    If (open_price = 0 And close_price = 0) Then
    percent_change = 0
    ElseIf (open_price = 0 And close_price <> 0) Then
    percent_change = 1
    Else
    percent_change = yearly_change / open_price
    Cells(summary_table_row, column + 11).Value = percent_change
    Cells(summary_table_row, column + 11).NumberFormat = "0.00%"
    
    End If
    
    ' Determine Total Volume
    
    total_volume = total_volume + Cells(i, column + 6).Value
    Cells(summary_table_row, column + 9).Value = total_volume
    summary_table_row = summary_table_row + 1
    
    open_price = Cells(i + 1, column + 2)
    
    total_volume = 0
    
    Else
    total_volume = total_volume + Cells(i, column + 6).Value
    
    End If
    
    Next i
    
    YCLastRow = ws.Cells(Rows.Count, column + 8).End(xlUp).Row
    
    ' Color code yearly change
    
    For j = 2 To YCLastRow
    If (Cells(j, column + 10).Value > 0 Or Cells(j, column + 10).Value = 0) Then
    Cells(j, column + 10).Interior.ColorIndex = 10
    ElseIf Cells(j, column + 10).Value < 0 Then
    Cells(j, column + 10).Interior.ColorIndex = 3
    
    End If
    
    Next j
    
    Next ws
    
    
    
    
End Sub
