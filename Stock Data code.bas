Attribute VB_Name = "Module1"
Sub Stocks()

'Initialize Variables
Dim Row As Long
Dim Col As Integer
Dim LastRowStock As Long
Dim LastRowCombined As Long
Dim AggRow As Long
Dim ws As Worksheet
Dim FirstRowTicker As Long
Dim LastRowTicker As Long
Dim RowVolume As LongLong
Dim TotalVolume As LongLong








AggRow = 2


'Add New Sheet for Combined Data

Sheets.Add.Name = "Combined Data"

'Add Column Headers to New Sheet

Range("A1").Value = "Ticker"
Range("B1").Value = "Yearly Change"
Range("C1").Value = "Percent Change"
Range("D1").Value = "Total Stock Volume"


'Create loop for first three columns of consolidated table


For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    FirstRowTicker = 2
    

    

    For Row = 2 To LastRow + 1
    

    
        If ws.Name <> "Combined Data" And ws.Cells(Row + 1, 1) <> ws.Cells(Row, 1) Then
        
        
        
        'Add Column Tickers
            'In each worksheet, for each change in the ticker from the first non-header row to the last row
            'Add the new ticker to column A of the new sheet
        
          Sheets("Combined Data").Cells(AggRow, 1) = ws.Cells(Row, 1)
          
          'Add Yearly Change
            'In each worksheet, for each change in the ticker from the first non header row to the last row
            'Calculate the difference in the open price in the first row of the ticker
            'And the last row of the ticker

          
          Sheets("Combined Data").Cells(AggRow, 2) = ws.Cells(Row, 6).Value - ws.Cells(FirstRowTicker, 3)
          
            If Sheets("Combined Data").Cells(AggRow, 2) <> 0 Then
          
                Sheets("Combined Data").Cells(AggRow, 3) = Sheets("Combined Data").Cells(AggRow, 2) / ws.Cells(Row, 6).Value
          
            Else
          
                Sheets("Combined Data").Cells(AggRow, 3) = 0
            
            End If
            
            
        'Format Percentages
        
        
          Sheets("Combined Data").Cells(AggRow, 3).NumberFormat = "0.00%"
          
          
        'Add Conditional Color Formatting to Yearly Change
          
          
          If Sheets("Combined Data").Cells(AggRow, 2) > 0 Then
        
                Sheets("Combined Data").Cells(AggRow, 2).Interior.ColorIndex = 4
                
          ElseIf Sheets("Combined Data").Cells(AggRow, 2) < 0 Then
          
                Sheets("Combined Data").Cells(AggRow, 2).Interior.ColorIndex = 3
                
          Else
          
          
          'Added if 0 make it yellow
          
              Sheets("Combined Data").Cells(AggRow, 2).Interior.ColorIndex = 6
              
          End If
          
          

          
                      
          FirstRowTicker = Row + 1
          
          
        AggRow = AggRow + 1
        
          
        End If
        
        
        

        
        
        
        
    Next Row
       
Next ws



'Total Volume Column Loop



AggRow = 2

For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For Row = 2 To LastRow + 1
    
        If ws.Name <> "Combined Data" And ws.Cells(Row + 1, 1) <> ws.Cells(Row, 1) Then
        
        'Set the Total Volume'
        
            TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
        
        'Print the Total Volume to Column 4 of the New Worksheet
        
            Sheets("Combined Data").Cells(AggRow, 4).Value = TotalVolume
        
        'Continue to Next Row on New Worksheet
            
            AggRow = AggRow + 1
        
        
        'Reset the Total Volume to 0
        
            TotalVolume = 0
            
            
        ElseIf ws.Name <> "Combined Data" And ws.Cells(Row + 1, 1) = ws.Cells(Row, 1) Then
        
        
            TotalVolume = TotalVolume + ws.Cells(Row, 7)
            
        
        End If
        
    Next Row

Next ws




End Sub

