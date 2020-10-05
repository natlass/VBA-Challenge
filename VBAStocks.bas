Sub VBAStocks()

'Loop through each worksheet
    Dim wb As Workbook
    Dim Main As Worksheet
         
' Set main worksheet
    Set wb = ActiveWorkbook
         
' Begin the loop
For Each Main In wb.Sheets

'Set headers on first worksheet
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

Next Main
    
For Each Main In Worksheets

    'Set variables
        Dim Ticker As String
            Ticker = " "
        Dim Percent_Change As Double
            Percent_Change = 0
        Dim Total_Stock_Volume As Double
            Total_Stock_Volume = 0
        Dim Closing_price As Double
            Closing_price = 0
        Dim Change_price As Double
            Change_price = 0
        Dim Open_price As Double
            Open_price = 0
        Dim Summary_Table As Long
            Summary_Table = 2
        Dim LastRow As Long
            LastRow = Main.Cells(Rows.Count, 1).End(xlUp).Row
     
     'Set beginning price for first ticker
     Open_price = Main.Cells(2, 3).Value
                 
    'Loop through all ticker changes
        For j = 2 To LastRow
            
        'Check if it is the same ticker, and enact then statement
            If Main.Cells(j + 1, 1).Value <> Main.Cells(j, 1).Value Then
            
            'State ticker value and spot for ticker within summary table
                Ticker = Main.Cells(j, 1).Value
                Main.Range("I" & Summary_Table).Value = Ticker
                    
            'State closing and opening price within loop
                Closing_price = Main.Cells(j, 6).Value
                                               
                Change_price = Closing_price - Open_price
                Main.Range("J" & Summary_Table).Value = Change_price
                    
            'Create Percent change and location within Summary Table with Formatting
                If Open_price <> 0 Then
                    Percent_Change = (Closing_price / Open_price) * 100
                End If
                
                Main.Range("K" & Summary_Table).Value = Percent_Change
            
            'Create total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Main.Cells(j, 7).Value
                    
            'Put the total stock volume in the summary table
                Main.Range("L" & Summary_Table).Value = Total_Stock_Volume
                    
            'Add 1 to summary table row count
                Summary_Table = Summary_Table + 1
              
            'Set conditional formatting
                If Change_price > 0 Then
                    Main.Range("J" & Summary_Table).Interior.ColorIndex = 4
                    
                ElseIf Change_price <= 0 Then
                    Main.Range("J" & Summary_Table).Interior.ColorIndex = 3
                
                End If
                    
            'Reset stock volume total to zero
                Total_Stock_Volume = 0
                Percent_Change = 0
                
               'For change in ticker name, enter new stock volume
            Else
            
                Total_Stock_Volume = Total_Stock_Volume + Main.Cells(j, 7).Value
                          
            End If
            
        Next j
        
Next Main
    
End Sub