# VBA_Challenge

Sub StockAnalysis()

'Data may have already been sorted, but sorted all data by Ticker and Date

'Declaring Variables

Dim New_Ticker As Boolean
Dim Ticker As String
Dim Total_Volume As LongLong
Dim Start_Price As Double
Dim End_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Output_Row As Long
Dim lastrow As Long

'Declaring Variables for Bonus

Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As LongLong
Dim ws As Worksheet
Dim GI_Ticker As String
Dim GD_Ticker As String
Dim GV_Ticker As String

'To run the code through each worksheet in the workbook.

For Each ws In Worksheets
    
    ws.Activate
    
    'Labeling Headers
    
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    
    'Reset the Bonus variables to 0 after finding the Bonus values for each worksheet.
    
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Volume = 0
    
    'New_Ticker set to true in order to read the first ticker.
    'Setting the first row to display the first ticker and all its calculations.
    'lastrow defined to find the last row to read for the last ticker.
    'Setting the format for the column displaying Percent Change as percents.
    
    New_Ticker = True
    Output_Row = 2
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("L2:L" & lastrow).NumberFormat = "0.00%"
    
    'Starting loop through all tickers in current worksheet.
    'Whenever a new ticker is found, save the value for the start price of that day
    'Most likely the first day of the year for that ticker.
    'New_Ticker now false since no longer the first entry of a new ticker,
    'remains until next ticker is found.

    For i = 2 To lastrow
    
        If New_Ticker = True Then
            Start_Price = Cells(i, 3).Value
            New_Ticker = False
        End If
        
        'The value in the first column in each row is set to the value in <ticker> column.
        'Take the value in volume column for each row and add together all volume values found.
    
        Ticker = Cells(i, 1).Value
        Total_Volume = Total_Volume + Cells(i, 7).Value
        
        'Recognizes when the ticker in the next row is not the same as the current value set in ticker.
        'When new ticker found, save the end price for that row, most likely the end price on the
        'last day of the year.  Calculate needed values, then print values in cells along current
        'row output.
    
        If Ticker <> Cells(i + 1, 1).Value Then
            If Start_Price > 0 Then
                End_Price = Cells(i, 6).Value
                Yearly_Change = End_Price - Start_Price
                Percent_Change = Yearly_Change / Start_Price
                Cells(Output_Row, 10).Value = Ticker
                Cells(Output_Row, 11).Value = Yearly_Change
                Cells(Output_Row, 12).Value = Percent_Change
                Cells(Output_Row, 13).Value = Total_Volume
                
                'Compares total volume for each ticker in order to find which is the greatest
                'among them.  Saves the value for volume if greater than previous and ticker
                'associated with that volume.
            
                If Total_Volume > Greatest_Volume Then
                    Greatest_Volume = Total_Volume
                    GV_Ticker = Ticker
                End If
                
                'Same as above if statement, but finds the greatest % increase and decrease
                'among all tickers.  Saves the value for percent change if greater than previous
                'or lesser than previous as well as tickers associated with each.
                
                If Percent_Change > Greatest_Increase Then
                    Greatest_Increase = Percent_Change
                    GI_Ticker = Ticker
                ElseIf Percent_Change < Greatest_Decrease Then
                    Greatest_Decrease = Percent_Change
                    GD_Ticker = Ticker
                End If
                
                'Checks if yearly change is greater or less than 0, colors the cell green
                ' for greater than 0 and red for less than 0.
            
                If Yearly_Change < 0 Then
                    Cells(Output_Row, 11).Interior.ColorIndex = 3
                Else
                    Cells(Output_Row, 11).Interior.ColorIndex = 4
                End If
                
                'Output row increases by one after each tickers values calculated to continue
                'displaying values for every ticker and their values consecutively.
            
                Output_Row = Output_Row + 1
            End If
            
            'New ticker is found, values will be calculated for a new ticker.
            'Resets volume to zero to add up all volumes for next ticker.
        
            New_Ticker = True
            Total_Volume = 0
        
        End If
        

    Next i
    
    'Displays all the bonus values calculated and saved for each worksheet.
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(2, 16).Value = GI_Ticker
    Cells(2, 17).Value = Greatest_Increase
    Range("Q2").NumberFormat = "0.00%"
    Cells(3, 16).Value = GD_Ticker
    Cells(3, 17).Value = Greatest_Decrease
    Range("Q3").NumberFormat = "0.00%"
    Cells(4, 16).Value = GV_Ticker
    Cells(4, 17).Value = Greatest_Volume


Next

End Sub
