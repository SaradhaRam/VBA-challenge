Sub multiple_year_stock_data():

  ' Set an initial variable for holding the ticker name
  Dim Ticker As String

  ' Set an initial variable for holding the total stock volume for each Ticker
  Dim Vol_Total As Double
  Vol_Total = 0

  ' Set an initial variable for holding the percentage of price Change
  Dim Percent_Change As Double
  Percent_Change = 0

  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Long

  ' Set an initial variable for holding the yearly price change
  Dim Yearly_Price_Change As Double

  Dim ws As Worksheet

  ' Set an initial variable for holding the lastrow
  Dim LastRow As Long

  ' Set a variable name ticker index
  Dim Ticker_Index As Long

    ' Loop through all the WorkSheets
    For Each ws In Worksheets

        ' Activate the worksheet
        ws.Activate
        
        ' Initialize the starting row of the summary table
        Summary_Table_Row = 2
        Ticker_Index = 2

        'Dynamically calculate the last row of the worksheet
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
             
             'Loop through year wise worksheet
             For I = 2 To LastRow
                
                ' Check if we are still within the same ticker name, if it is not...
                If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                    
                    ' Set the ticker name
                    Ticker = Cells(I, 1).Value
                    
                    ' Add to the stock volume Total
                    Vol_Total = Vol_Total + Cells(I, 7).Value

                    ' Calculate price change by subracting the closing price and opening price
                    Yearly_Price_Change = Cells(I, 6).Value - Cells(Ticker_Index, 3).Value


                        ' Check if the opening price is not equal to zero and not null to avoid divide by zero error
                         If (Cells(Ticker_Index, 3).Value <> 0 And Not IsNull(Cells(Ticker_Index, 3).Value)) Then

                                'Percent change calculation done by divide the yearly price change by opening price
                                 Percent_Change = (Yearly_Price_Change / Cells(Ticker_Index, 3).Value)

                         End If

                    ' Writing the Calculated values into summary table
                    Range("I" & Summary_Table_Row).Value = Ticker
                    Range("J" & Summary_Table_Row).Value = Vol_Total
                    Range("K" & Summary_Table_Row).Value = Yearly_Price_Change
                    Range("L" & Summary_Table_Row).Value = Percent_Change

                    Summary_Table_Row = Summary_Table_Row + 1
                    Vol_Total = 0
                    Ticker_Index = I + 1
    
                Else

                    Vol_Total = Vol_Total + Cells(I, 7).Value

                End If
            
            Next I

    Next ws
    
            '******* Code for conditional fomatting the Yearly price change column*******
    
    ' Loop through each worksheet for Conditional formating and calculate Min and Max values
    For Each ws In Worksheets

        Dim LastRow1 As Long
        ws.Activate
        
        ' Calculate the last row dynamically for the percent change column(K)
        LastRow1 = Cells(Rows.Count, 11).End(xlUp).Row
            
            ' Loop through each row if less than zero color the cell red if not color green
            For I = 2 To LastRow1

                If Cells(I, 11) < 0 Then
                    Cells(I, 11).Interior.ColorIndex = 3  ' 3 is Red
                Else
                    Cells(I, 11).Interior.ColorIndex = 4  ' 4 is Green
                End If

            Next I
        
            '********** Code for Greatest % Increase / Greatest % decrease / Greatest total Volume**********
            
        ' Set the variable name for Range and Max,Min percentage and Max total volume
        Dim rng, rng1 As Range
        Dim max_percent_change, min_percent_change As Double
        Dim max_total_volume As LongLong

        ' Set range from which to determine Max/Min value
        Set rng = Range("L2:L" & LastRow1)
        Set rng1 = Range("J2:J" & LastRow1)

        ' Worksheet function MAX/MIN returns the largest/smalest value in a range
        max_percent_change = Application.WorksheetFunction.Max(rng)
        min_percent_change = Application.WorksheetFunction.Min(rng)
        max_total_volume = Application.WorksheetFunction.Max(rng1)

        ' Diaplay the names in the given cell values.
        Cells(5, 14).Value = "Greatest % Increase"
        Cells(6, 14).Value = "Greatest % Decrease"
        Cells(7, 14).Value = "Greatest Total Volume"

        ' Display the values in the respective names.
        Cells(5, 15).Value = max_percent_change
        Cells(6, 15).Value = min_percent_change
        Cells(7, 15).Value = max_total_volume

    Next ws

End Sub




