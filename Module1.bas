Attribute VB_Name = "Module1"
Sub alphabetical_testing():
Dim ws As Worksheet
For Each ws In Worksheets
    ' Set dimensions
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
         
    
     'set variables
    Dim Ticker As String
    Dim open_price As Double
    open_price = ws.Cells(2, 3).Value
    
    Dim closing_price As Double
    closing_price = 0
     
    output_row = 2
  
    


 
    'set a variable for lrow
    Dim lrow As Long
 
    'Find the last cell column A
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 

    'start for loop
      
    For i = 2 To lrow
        ' Stores results in variables
        Ticker = ws.Cells(i, 1).Value
        Total = Total + ws.Cells(i, 7).Value

       

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            closing_price = ws.Cells(i, 6).Value
 
            'calculate change
            YearlyChange = closing_price - open_price
  
            'calculate percent change
            PercentChange = (YearlyChange) / (open_price)

            ws.Cells(output_row, 9) = Ticker
            ws.Cells(output_row, 10) = YearlyChange
            ws.Cells(output_row, 11) = PercentChange
            ws.Cells(output_row, 12) = Total

            
            
            
            open_price = ws.Cells(i + 1, 3).Value
            output_row = output_row + 1
            Total = 0
            
        End If
                
    Next i
    
    DataStart = 2
    DataEnd = ws.Cells(Rows.Count, 10).End(xlUp).Row
    For j = DataStart To DataEnd
    
        If ws.Cells(j, 10) > 0 Then
            ws.Cells(j, 10).Interior.Color = vbGreen
        Else
            ws.Cells(j, 10).Interior.Color = vbRed
        End If
    Next j
    
Next ws
        
End Sub

