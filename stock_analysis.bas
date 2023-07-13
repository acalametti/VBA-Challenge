Attribute VB_Name = "Module1"

Sub stock_analysis2():
  
  'Dim ws as Worksheet
  
  'For Each ws In Workbook (from class activity)
     'I couldn't get the ws loop to work without changing the correct values in the spreadsheet so I just added what I had tried
        'I know ws is supposed to be added in front of any Cells(x,y) statements
     

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row 'from class activity

     
     ''DEFINE ALL VARAIBLES
        'TICKER
        Dim Ticker As String
        
        'VARIABLES for YEARLY CHANGE
        Dim Open_amount As Double
        Dim Close_amount As Double
        Dim yearly_change As Double
        yearly_change = 0
        
        'PERCENT CHANGE
        Dim percent_change As Double
        percent_change = 0
        
        'TOTAL STOCK
        Dim total_stock As Double
             
        total_stock = 0
        
        'SUMMARY TABLE
        Dim summary_row As Integer
        summary_row = 2
        s = 2 ' from learning assistant on BCS learning (Dinh)
        
    'START FOR LOOP
    For i = 2 To 753001
     
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'from class activity
            
            'IDENTIFY TICKER CELLS
            Ticker = Cells(i, 1).Value
            
            'PRINT TICKER IN SUMMARY TABLE
            Range("I" & summary_row).Value = Ticker
            
            'ADD TO STOCK VOLUME
            total_stock = total_stock + Cells(i, 7).Value
            
            'ADD STOCK VOLUME TO TABLE
            Range("L" & summary_row).Value = total_stock
            
            'ADD THE CLOSING AMOUNT
            Close_amount = Cells(i, 6).Value
            
            'ADD OPEN AMOUNT ' from learning assitant
            Open_amount = Cells(s, 3).Value
    
            'CALCULATE THE YEARLY CHANGE
            yearly_change = Close_amount - Open_amount
        
            'ASSIGN SPOT IN SUMMARY TABLE
            Range("J" & summary_row).Value = yearly_change
            
            'CALCULATE THE PERCENT CHANGE
            percent_change = (yearly_change / Open_amount)
                    
             'ADD PERCENT CHANGE TO SUMMARY TABLE
            Range("K" & summary_row).Value = percent_change
            
            'HAVE SUMMARY ROW FILL NEXT ROW
            summary_row = summary_row + 1
            
           'RESET STOCK VOLUME
            total_stock = 0
            yearly_change = 0
            s = i + 1
                  
         Else
            'ADD TO STOCK TOTAL
            total_stock = total_stock + Cells(i, 7).Value
                
                           
              
     End If
            

    Next i
'Next ws

End Sub


