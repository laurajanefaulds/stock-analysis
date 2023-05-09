Sub stock_analysis()

    For Each ws In Worksheets
    
        ' Set an initial variable for holding the brand name
        Dim Ticker_Name As String

        ' Set an initial variable for holding the total per ticker
        Dim Vol_Total As Double
        Vol_Total = 0

        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        'Define Variables
        Dim Close_Price As Double
        Close_Price = 0

        Dim Start_Price As Double
        Start_Price = ws.Cells(2, 3).Value

        Dim Yearly_Change As Double
        Yearly_Change = 0

        Dim Percent_Change As Double
        Percent_Change = 0
        
        'Format Percent Changed Column as %
        Range("K:K").NumberFormat = "0.00%"

        ' Loop through all ticker transactions
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' Check if we are still within the same ticker type, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker name
                Ticker_Name = ws.Cells(i, 1).Value

                ' Add to the Volume Total
                Vol_Total = Vol_Total + ws.Cells(i, 7).Value

                ' Set the close price value
                Close_Price = ws.Cells(i, 6).Value

                ' Print the Ticker Name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                ' Print the Total Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Vol_Total

                'Set new value for Yearly Change
                Yearly_Change = Close_Price - Start_Price

                'Print the Yearly Change Value in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Set new value for Percent Change
                Percent_Change = Yearly_Change / Start_Price

                'Apply conditional formatting to Yearly Change column
                If Yearly_Change < 0 Then
                    'Set cell background colour to red
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3

                Else
                    'Set cell background colour to green
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4

                End If

                'Print the Percent Changed Value in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change

                ' Reset the Volume Total and Start Price for the next ticker type
                Summary_Table_Row = Summary_Table_Row + 1

                ' Reset the Volume Total for the next ticker type
                Vol_Total = 0

                Start_Price = ws.Cells(i + 1, 3).Value

            Else
                ' If we are still within the same ticker type, add to the Volume Total
                Vol_Total = Vol_Total + ws.Cells(i, 7).Value
            End If
        Next i

    Next ws
    
    'Across all worksheets
    For Each ws In Worksheets
    
        'Define Variables for New Table
    Dim GPInc_Ticker As String
    GPInc_Ticker = ""
    
    Dim GPInc_Value As Double
    GPInc_Value = 0
    
    Dim GPDec_Ticker As String
    GPDec_Ticker = ""
    
    Dim GPDec_Value As Double
    GPDec_Value = 0
    
    Dim GTVol_Ticker As String
    GTVol_Ticker = ""
    
    Dim GTVol_Value As Double
    GTVol_Value = 0
    
    'Loop through all relevant column
    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
         
        'Check for greatest % increase
         If Cells(i, 11).Value >= GPInc_Value Then
            GPInc_Value = ws.Cells(i, 11).Value
            GPInc_Ticker = ws.Cells(i, 9).Value
            
        End If
            
            'Check for Greatest % Decrease
        If Cells(i, 11).Value <= GDec_Value Then
            GPDec_Value = ws.Cells(i, 11).Value
            GPDec_Ticker = ws.Cells(i, 9).Value
        
        End If
        
            'Check for Greatest Total Volume
          If Cells(i, 12).Value >= GTVol_Value Then
            GTVol_Value = ws.Cells(i, 12).Value
            GTVol_Ticker = Cells(i, 9).Value
        
        End If
        
        Next i
    
      ' Add the word Ticker to the Column I Header
        ws.Cells(1, 9).Value = "Ticker"
        
        ' Add the words Yearly Change to the Column J Header
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ' Add the words Percent Change to the Column K Header
        ws.Cells(1, 11).Value = "Percent Change"
        
        ' Add the words Total Stock Volume to the Column L Header
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Add the words Greatest % Increase to Cell O2
        ws.Cells(2, 15).Value = "Greatest % Increase"
        
          ' Add the words Greatest % Decrease to Cell O3
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        
          ' Add the words Greatest Total Volume to Cell O4
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Add the word Ticker to the Column P Header
        ws.Cells(1, 16).Value = "Ticker"
        
        ' Add the word Value to the Column Q Header
        ws.Cells(1, 17).Value = "Value"
    
    'Assign values to corresponding cell in table
    ws.Range("P2") = GPInc_Ticker
    ws.Range("Q2") = GPInc_Value
    
     ws.Range("Q2").NumberFormat = "0.00%"

    ws.Range("P3") = GPDec_Ticker
    ws.Range("Q3") = GPDec_Value
    
     ws.Range("Q3").NumberFormat = "0.00%"

    ws.Range("P4") = GTVol_Ticker
    ws.Range("Q4") = GTVol_Value
    
        Next ws
    
End Sub