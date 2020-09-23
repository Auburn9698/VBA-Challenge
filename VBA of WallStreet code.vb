Sub Stock_Counter()

For Each ws In Worksheets
ws.Activate
    ' Set an initial variable for the Ticker Name:
    Dim Ticker_Name As String
    
    ' Set an initial variable for opening price, closing price, and yearly change:
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    ' Set initial opening price:
    opening_price = Cells(2, 3).Value
    
    ' Set variable for holding total volume, starting at 0:
    Dim Volume_Total As Double
    Volume_Total = 0
    
    ' Determine the last row:
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Define Summary Row Table, set initially:
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Loop through all the stock information:
    For i = 2 To LastRow
    
        ' Check to see if we are still within the same stock, and if we're not:
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Define the Ticker Name:
            Ticker_Name = Cells(i, 1).Value
            
            ' Print the Ticker_Name in the Summary Table:
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            
            ' Define Close Price
            closing_price = Cells(i, 6).Value
            
            ' Define Yearly Change
            yearly_change = closing_price - opening_price
            
            ' Print the yearly change to the summary table:
            Range("J" & Summary_Table_Row).Value = yearly_change
            
            'Error catcher for prices of 0 (...not working?...):
            If yearly_change = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / opening_price
            
                ' Print the percent change to the summary table:
            Range("K" & Summary_Table_Row).Value = percent_change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
            End If
            
            ' Add to the Volume Total:
            Volume_Total = Volume_Total + Cells(i, 7).Value
            
            ' Print the Volume Total to the Summary Table:
            Range("L" & Summary_Table_Row) = Volume_Total
     
            ' Add one to the Summary Table Row:
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the Volume_Total
            Volume_Total = 0
            
            ' Reset the opening price:
            opening_price = Cells(i + 1, 3).Value
           
        Else
            
            ' If the next cell IS the same stock:
            Volume_Total = Volume_Total + Cells(i, 7).Value
                
        End If
        
    Next i
    
    ' Determine last row of yearly change column:
    YC_LastRow = Cells(Rows.Count, 10).End(xlUp).Row
    
    'Format red or green for positive or negative
    For j = 2 To YC_LastRow
        If (Cells(j, 10).Value >= 0) Then
            Cells(j, 10).Interior.ColorIndex = 4
        Else
            Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j
    
    ' Find Greatest % Increase or Decrease:
    For k = 2 To YC_LastRow
        If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & YC_LastRow)) Then
            Cells(2, 15).Value = Cells(k, 9).Value
            Cells(2, 16).Value = Cells(k, 11).Value
            Cells(2, 16).NumberFormat = "0.00%"
        ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & YC_LastRow)) Then
            Cells(3, 15).Value = Cells(k, 9).Value
            Cells(3, 16).Value = Cells(k, 11).Value
            Cells(3, 16).NumberFormat = "0.00%"
        End If
    Next k
    
    ' Find Greatest Total Volume:
    For m = 2 To YC_LastRow
        If Cells(m, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & YC_LastRow)) Then
            Cells(4, 15).Value = Cells(m, 9).Value
            Cells(4, 16).Value = Cells(m, 12).Value
        End If
    Next m
    
  
            
    ' Set headers for the summary tables:
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"

    
    'Autofit  sheets:
    Application.ScreenUpdating = False
    Dim wkSt As String
    Dim wkBk As Worksheet
    wkSt = ActiveSheet.Name
    For Each wkBk In ActiveWorkbook.Worksheets
        On Error Resume Next
        wkBk.Activate
        Cells.EntireColumn.AutoFit
    Next wkBk
    Sheets(wkSt).Select
    Application.ScreenUpdating = True
        

Next ws

End Sub
