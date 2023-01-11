'****************************Multi-Year Stock Data*****************************************************************************
'Instructions
'Create a script that loops through all the stocks for one year and outputs the following information:
    'The ticker symbol
    'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock. The result should match the following image:
    'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    'The solution should match the following image:
    'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
'Note:
    'Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
'Other Considerations:
    'Use the sheet alphabetical_testing.xlsx while developing your code.
    'This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.
    'Make sure that the script acts the same on every sheet.
    'The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

Sub Multi_year_stock_data()


'Looping through each Worksheet in the excel
    For Each Ws In Worksheets
        
        'Worksheetname
        Dim Worksheetname As String
        
        'Current row of Stock
        Dim Curr_Row As Long
        'Start row of ticker set
        Dim Start_Row As Long
        'Index counter to fill Ticker row
        Dim Tickr_Cnt As Long
        'Variable for percent change calculation
        Dim Pct_Chng As Double
        'Last row column of Stock
        Dim Last_Row_Stock As Long
    
        'last row column Summary
        Dim Last_Row_Summary As Long
        'Variable for greatest increase calculation
        Dim Grt_Incr As Double
        'Variable for greatest decrease calculation
        Dim Grt_Dcr As Double
        'Variable for greatest total volume
        Dim Grt_Vol As Double
        
        'Get the WorksheetName
        Worksheetname = Ws.Name
   
   
   '****** Summary of Stock Values*****
        
        
        'Populate column headers
        Ws.Cells(1, 10).Value = "Ticker"
        Ws.Cells(1, 11).Value = "Yearly Change"
        Ws.Cells(1, 12).Value = "Percent Change"
        Ws.Cells(1, 13).Value = "Total Stock Volume"
        
        'Set Ticker Counter to first row of data
        Tickr_Cnt = 2
        
        'Set start row to thethe first row of data
        Start_Row = 2
        
        'Find the number of rows that has values
        Last_Row_Stock = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows that has stock data
            For Curr_Row = 2 To Last_Row_Stock
            
                'condition to check the ticker block
                If Ws.Cells(Curr_Row + 1, 1).Value <> Ws.Cells(Curr_Row, 1).Value Then
                
                    'Once we reach a different ticker set store the summary of the previous one
                    
                    'Ticker name
                    Ws.Cells(Tickr_Cnt, 10).Value = Ws.Cells(Curr_Row, 1).Value
                
                     'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
                     Ws.Cells(Tickr_Cnt, 11).Value = Ws.Cells(Curr_Row, 6).Value - Ws.Cells(Start_Row, 3).Value
                
                    'Conditional formating for Yearly Change column
                    If Ws.Cells(Tickr_Cnt, 11).Value < 0 Then
                        
                        'Set color to red if value is less than 0
                        Ws.Cells(Tickr_Cnt, 11).Interior.ColorIndex = 3
                    
                    Else
                        
                        'Set color to green if value is greater than or equal to 0
                        Ws.Cells(Tickr_Cnt, 11).Interior.ColorIndex = 4
                    
                    End If
                    
                    'Percentage Change
                    If Ws.Cells(Start_Row, 3).Value <> 0 Then
                        
                        'calculate the Percentage change value
                        Pct_Chng = ((Ws.Cells(Curr_Row, 6).Value - Ws.Cells(Start_Row, 3).Value) / Ws.Cells(Start_Row, 3).Value)
                        
                       'Formatting the cell when it has value value
                        Ws.Cells(Tickr_Cnt, 12).Value = Format(Pct_Chng, "Percent")
                    
                    Else
                        
                        Ws.Cells(Tickr_Cnt, 12).Value = Format(0, "Percent")
                    
                    End If
                    
                    'Total Stock Count - Sum from first to last row in the ticker set
                    Ws.Cells(Tickr_Cnt, 13).Value = WorksheetFunction.Sum(Range(Ws.Cells(Start_Row, 7), Ws.Cells(Curr_Row, 7)))
                
               'Increment Ticker Count
                Tickr_Cnt = Tickr_Cnt + 1
                
                'Set new start row of the ticker block
                Start_Row = Curr_Row + 1
                
                End If
            
            Next Curr_Row
            

            
'************Greatest Values*******************

        'Populate Column Headers
        Ws.Cells(1, 17).Value = "Ticker"
        Ws.Cells(1, 18).Value = "Value"
        Ws.Cells(2, 16).Value = "Greatest % Increase"
        Ws.Cells(3, 16).Value = "Greatest % Decrease"
        Ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        'identify the number of rows in the Summary section
        Last_Row_Summary = Ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        
        'Set the Values initially to the first row data value in the Summary section
        Grt_Incr = Ws.Cells(2, 12).Value
        Grt_Decr = Ws.Cells(2, 12).Value
        Grt_Vol = Ws.Cells(2, 13).Value
        
        
            'Loop through the Summary Section
            For Curr_Row = 2 To Last_Row_Summary
                
                
                'Greatest Increse - Compare the current row value with the next row value.
                If Ws.Cells(Curr_Row, 12).Value > Grt_Incr Then
                    
                    'If its greater then store it to the variable
                    Grt_Incr = Ws.Cells(Curr_Row, 12).Value
                    
                    'Populate the Greatest section with the ticker value & the corresponding Percentage Change with formatting
                    Ws.Cells(2, 17).Value = Ws.Cells(Curr_Row, 10).Value
                    Ws.Cells(2, 18).Value = Format(Grt_Incr, "Percent")
                
                Else
                
                    Grt_Incr = Grt_Incr
                
                End If
                
                'Greatest Decrease - Compare current row value with the next row value
                If Ws.Cells(Curr_Row, 12).Value < Grt_Decr Then
                    
                    'If its smaller then store it to the variable
                    Grt_Decr = Ws.Cells(Curr_Row, 12).Value
                    
                   'Populate the Greatest section with the ticker value & the corresponding Percentage Change with formatting
                    Ws.Cells(3, 17).Value = Ws.Cells(Curr_Row, 10).Value
                    Ws.Cells(3, 18).Value = Format(Grt_Decr, "Percent")
                
                Else
                
                Grt_Decr = Grt_Decr
                
                End If
                
                'Greatest Volume - Compare current row value with the next row value
                If Ws.Cells(Curr_Row, 13).Value > Grt_Vol Then
                    
                    'If its greater then store it to the variable
                    Grt_Vol = Ws.Cells(Curr_Row, 13).Value
                    
                    'Populate the Greatest section with the ticker value & the corresponding Volume with formatting
                    Ws.Cells(4, 17).Value = Ws.Cells(Curr_Row, 10).Value
                    Ws.Cells(4, 18).Value = Format(Grt_Vol, "Scientific")
                
                Else
                
                Grt_Vol = Grt_Vol
                
                End If
                
    
            Next Curr_Row
            
        'Autofit the columns in the Worksheet
        Worksheets(Worksheetname).Columns("A:R").AutoFit
            
    Next Ws
        

End Sub
