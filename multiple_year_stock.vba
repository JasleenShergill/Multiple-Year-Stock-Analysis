Attribute VB_Name = "Module1"
Sub multiple_stock_data()

'INSTRUCTIONS
'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock. The result should match the following image:
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
'-----------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------

'applying code on each worksheet
For Each ws In Worksheets

'using ws.activate to eliminate the re-using ws with each command
ws.Activate

'declaring variables
Dim Worksheet As String
Dim percent_change, greatest_increase, greatest_decrease, greatest_vol As Double
Dim i, lastrow, lastrow_I As Long
Dim j As Long
Dim ticker_count As Long
' long was used due to integer not being able to hold the values , error : overflow

'assigning values to variables

ticker_count = 2

j = 2


'assigning headers to columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase "
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'To get worksheet name
worksheetname = ws.Name

'calculating last row of each worksheet
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
' for testing purpose run : MsgBox (lastrow)

'using loop function to go through each row
For i = 2 To lastrow

    ' checking to see if ticker changes
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'enter ticker value in column I
         Cells(ticker_count, 9).Value = Cells(i, 1).Value
    
         'calculating yearly changes
         Cells(ticker_count, 10).Value = Cells(i, 6).Value - Cells(j, 3).Value
        
      'conditional formating
      If Cells(ticker_count, 10).Value < 0 Then
        
            'set cell background to red ,using color index 3
            Cells(ticker_count, 10).Interior.ColorIndex = 3
        
        Else
        
            'set cell background to green, using color index 4
            Cells(ticker_count, 10).Interior.ColorIndex = 4
        
        End If
        'MsgBox ("working")
        
        'Calculate and write percent change in column K (#11)
        If Cells(j, 3).Value <> 0 Then
             percent_change = ((Cells(i, 6).Value - Cells(j, 3).Value) / Cells(j, 3).Value)
                    
            'Percent formating
             Cells(ticker_count, 11).Value = Format(percent_change, "Percent")
                    
        Else
                    
            Cells(ticker_count, 11).Value = Format(0, "Percent")
                    
        End If
    
    'calculate and write total volume in column L
     Cells(ticker_count, 12).Value = WorksheetFunction.Sum(Range(Cells(j, 7), Cells(i, 7)))
     
        
    'increase ticker counter by 1
    ticker_count = ticker_count + 1

    'move to next row
    j = i + 1
    
   End If
    
Next i

'After the 1st loop is done, set the next loop
      
       greatest_increase = Range("K2").Value
       greatest_decrease = Range("K2").Value
       greatest_vol = Range("L2").Value
       
       'finding last value of column I
       lastrow_I = Cells(Rows.Count, 10).End(xlUp).Row
  
       For i = 2 To lastrow_I
        
        'if next value is more than current value replace value with larger number
        If Cells(i, 11) > greatest_increase Then
        greatest_increase = Cells(i, 11)
        Range("Q2").Value = Format(greatest_increase, "Percent")
        Range("P2").Value = Cells(i, 9)
   
        End If
   
        'if next value is less than current value replace value with lower number
        If Cells(i, 11) < greatest_decrease Then
        greatest_decrease = Cells(i, 11)
        Range("Q3").Value = Format(greatest_decrease, "Percent")
        Range("P3").Value = Cells(i, 9)
  
        End If
        
        'if next value is more than current value replace value with larger number
        If Cells(i, 12) > greatest_vol Then
        greatest_vol = Cells(i, 12)
        Range("Q4").Value = Format(greatest_vol, "Scientific")
        Range("P4").Value = Cells(i, 9)
   
        End If
   
       Next i
 
        'formatting cells to autofit
        Columns("I:Q").AutoFit
        Range("A1:Q1").Font.Bold = True

Next ws

End Sub



