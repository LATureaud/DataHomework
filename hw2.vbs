Sub TickerLoop()

'Set initial variables for items needed
Dim ticker As String
Dim volume_total As Double
    volume_total = 0


Dim lastRow As Long
  
  
    'Keep track of location for each credit card brand in summary table
    Dim Summary_Table As Integer
    Summary_Table = 2
    
    
      
     
     For Each ws In Worksheets
        
        
    'For understanding lastRow = Last Row
   
   lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   Summary_Table = 2
    
    'Print names for titles
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
    ws.Range("J1").EntireColumn.AutoFit
     
     'Loop through all ticker data
         For i = 2 To lastRow

       
       'Check to see if we are still in the same ticker name, if not
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
       
       'Set the ticker name
       ticker = ws.Cells(i, 1).Value
       
       'Add to the volume total
       volume_total = volume_total + ws.Cells(i, 7).Value
       
       'Print the Ticker in the Summary Table
       ws.Range("I" & Summary_Table).Value = ticker
       
       'Print the volume total amount to the Summary Table
        ws.Range("J" & Summary_Table).Value = volume_total
       
       'Add one to the summary_table
       Summary_Table = Summary_Table + 1
       
       'Reset the volume_total
       volume_total = 0
       
       'If the cell immediately following a row is the same ticker
       Else
       
       'Add to the volume_total
           volume_total = volume_total + ws.Cells(i, 7).Value
       
       End If
    
    Next i
    
  Next ws
       
End Sub

