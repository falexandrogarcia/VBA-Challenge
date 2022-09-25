Attribute VB_Name = "Module1"
Sub VBA_Homework()

' Create titles
    [I1].Value = "Ticker"
    [J1].Value = "Yearly Change"
    [K1].Value = "Percent Change"
    [L1].Value = "Total Stock Volume"
    
    [I1:L1].Select
    Selection.Columns.AutoFit
    
' Yearly change from opening price at the beginning of a given
' year to the closing price at the end of that year
    
    ' Set a varible to indicate the last row number on a data set
    Dim Irow As Variant
    Irow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set an initial variable for holding the ticker name
    Dim Ticker_Name As String

    ' Set an initial variable for holding the total volume per ticker
    Dim Ticker_Total As Double
    Ticker_Total = 0

    ' Keep track of the location for each ticker in the
    ' summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Create a variable to store the stock opening and closing price
    Dim opening_price As Variant
    Dim closing_price As Variant
    opening_price = [C2].Value
    closing_price = 0

    ' Loop through all tickers data points
    For i = 2 To Irow

        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the Ticker name
            Ticker_Name = Cells(i, 1).Value

            ' Add the last row of the same ticker to the Ticker Total
            Ticker_Total = Ticker_Total + Cells(i, 7).Value
            
            ' Capture the closing stock price
            closing_price = Cells(i, 6).Value
            
            
            
            ' Create Summary Table ---------------------------
            ' Print the Ticker name in the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker_Name

            ' Print the Ticker Volume to the Summary Table
            Range("L" & Summary_Table_Row).Value = Ticker_Total
            
            ' Print the yearly change to the Summary Table
            Range("J" & Summary_Table_Row).Value = closing_price - opening_price
            
            ' Print the Percent Chage of the stock price for each ticker
            Range("K" & Summary_Table_Row).Value = FormatPercent((closing_price - opening_price) / opening_price, 2)
            
            ' Color format each cell based upon its value
            If Range("J" & Summary_Table_Row).Value < 0 Then
                
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            Else
                
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
            End If
                        

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the Ticker Total
            Ticker_Total = 0
            
            ' Reset opening stock price
            opening_price = Cells(i + 1, 3).Value

        ' If the cell immediately following a row is the same ticker...
        Else

            ' Add to the Ticker Total
            Ticker_Total = Ticker_Total + Cells(i, 7).Value

        End If

    Next i

End Sub
