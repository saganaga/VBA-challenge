Sub Multiple_year_stock_data()

    'Run on every worksheet
    For Each ws In Worksheets

    'Set value in I1 to Ticker
    ws.Cells(1, 9).Value = "Ticker"

    'Set value in J1 to Yearly Change
    ws.Cells(1, 10).Value = "Yearly Change"

    'Set value in K1 to Percent Change
    ws.Cells(1, 11).Value = "Percent Change"

    'Set value inL1 to Total Stock Volume
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Set an initial value for holding the Ticker
    Dim Ticker As String

    'Determine the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Set an initial value for holding the Yearly Change
    Dim Yearly_Change As Double

    'Set an initial value for holding the Percent Change
    Dim Percent_Change As Double

    'Set an initial value for holding the Total Stock Volume
    Dim Total_Stock_Volume As LongLong
    Total_Stock_Volume = 0

    'Keep track of the location of each Ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Keep track of current Ticker start row number
    Dim Start As Long
    Start = 2

    'Loop through all the Tickers
    For i = 2 To LastRow

        'Check to see if still within same Ticker, if it is not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set the Ticker
            Ticker = ws.Cells(i, 1).Value
            
            'Compute Yearly Change
            Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value
            
            'Compute Percent Change
            Percent_Change = Yearly_Change / ws.Cells(Start, 3).Value

            'Print Ticker in summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            'Print Yearly Change in summary table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
                'Highlight negative change in red and positive change in green
                If Yearly_Change < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
            
            'Print Percent Change in summary table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

            'Print Total Stock Volume to summary table
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

            'Reset the Total Stock Volume
            Total_Stock_Volume = 0
            
            'Reset Start for next Ticker
            Start = i + 1

        'If cell immediately down a row is same Ticker
        Else

        'Add to the Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

        End If

    Next i

    'Retain the index of the last row of the summary table
    Dim Last_Summary_Table_Row As Integer
    Last_Summary_Table_Row = Summary_Table_Row - 1

    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As LongLong
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volume = 0
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Total_Volume_Ticker As String


    'Loop through summary table
    For i = 2 To Last_Summary_Table_Row
        Ticker = ws.Cells(i, 9).Value

        'Find the max and min in Percent Change column
        Percent_Change = ws.Cells(i, 11).Value
        If Percent_Change > 0 Then
            'Check Greatest Increase
            If Percent_Change > Greatest_Increase Then
                Greatest_Increase = Percent_Change
                'Grab Greatest_Increase_Ticker
                Greatest_Increase_Ticker = Ticker
            End If
        Else
            'Check Greatest Decrease
            If Percent_Change < Greatest_Decrease Then
                Greatest_Decrease = Percent_Change
                'Grab Greatest_Decrease_Ticker
                Greatest_Decrease_Ticker = Ticker
            End If
        
        End If
        
        'Find max in Total Stock Volume column
        Total_Stock_Volume = ws.Cells(i, 12).Value
        If Total_Stock_Volume > Greatest_Total_Volume Then
            Greatest_Total_Volume = Total_Stock_Volume
            'Grab Greatest_Total_Volume_Ticker
            Greatest_Total_Volume_Ticker = Ticker
        End If
        
    Next i

    'Set value in O2 to Greatest % Increase
    ws.Cells(2, 15).Value = "Greatest % Increase"

    'Set value in O3 to Greatest % Decreases
    ws.Cells(3, 15).Value = "Greatest % Decrease"

    'Set value in O4 to Greatest Total Volume
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    'Set value in P1 to Ticker
    ws.Cells(1, 16).Value = "Ticker"

    'Set value in Q1 to Value
    ws.Cells(1, 17).Value = "Value"

    'Print Greatest_Increase_Ticker in (2, 16)
    ws.Cells(2, 16).Value = Greatest_Increase_Ticker
    'Print Greatest_Increase in (2, 17)
    ws.Cells(2, 17).Value = Greatest_Increase
    ws.Range("Q2").NumberFormat = "0.00%"
    'Print Greatest_Decrease_Ticker in (3, 16)
    ws.Cells(3, 16).Value = Greatest_Decrease_Ticker
    'Print Greatest_Decrease in (3, 17)
    ws.Cells(3, 17).Value = Greatest_Decrease
    ws.Range("Q3").NumberFormat = "0.00%"
    'Print Greatest_Total_Volume_Ticker in (4, 16)
    ws.Cells(4, 16).Value = Greatest_Total_Volume_Ticker
    'Print Greatest_Total_Volume in (4, 17)
    ws.Cells(4, 17).Value = Greatest_Total_Volume

    Next ws

End Sub

