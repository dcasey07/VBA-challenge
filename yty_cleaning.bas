Attribute VB_Name = "Module1"
Option Explicit

Sub yty_cleaning()
    Dim ws As Worksheet
    For Each ws In Worksheets

        ' Set an initial variable for holding the ticker symbol
        Dim ticker_sym As String

        ' Set an initial variable for holding the total stock volume
        Dim stock_vol As Double
            stock_vol = 0
        
        ' Set initial variables for tracking the stock with the greatest percentage increase
        Dim ticker_greatest_incr As String
        Dim greatest_incr As Double
            greatest_incr = 0
        ' Same for stock with greatest percentage decrease
        Dim ticker_greatest_decr As String
        Dim greatest_decr As Double
            greatest_decr = 0
        ' Same for stock with greatest total stock volume
        Dim greatest_vol_ticker As String
        Dim greatest_vol As Double
            greatest_vol = 0
        
        ' Set an initial variable for determining the last row in the data set
        Dim last_row As Long
        
        ' Set an initial variable for determining the opening stock price for the year
        Dim open_stock As Double
        
        ' Same but for closing stock price on the year
        Dim close_stock As Double
        
        ' Keep track of the location for each ticker symbol in the yearly output table where the data starts
        Dim yearly_output As Long
            yearly_output = 2
        
        ' Set a value for the last row in the worksheet
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set an advance integer for the loop and start it on the row where the data starts
        Dim i As Long
            i = 2
        
        ' Set the value of the opening stock value
        open_stock = ws.Cells(i, 3).Value
        
        ' Output the opening stock value
        ws.Range("J" & yearly_output).Value = open_stock
        
        ' Loop through all ticker values
        For i = 2 To last_row
            ' Run a check for when the ticker name changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker symbol
                ticker_sym = ws.Cells(i, 1).Value
                ' Extract the stock value in the set
                close_stock = ws.Cells(i, 6).Value
                ' Add to the total stock volume
                stock_vol = stock_vol + ws.Cells(i, 7).Value
                ' Calculate and output the yearly change
                ws.Range("J" & yearly_output).Value = close_stock - open_stock
                ' Calculate and output the yearly percentage change
                ws.Range("K" & yearly_output).Value = ((close_stock - open_stock) / open_stock)
                ' Output the ticker symbol in the yearly output table
                ws.Range("I" & yearly_output).Value = ticker_sym
                ' Output the total stock volume to the yearly output table
                ws.Range("L" & yearly_output).Value = stock_vol
                ' Check for greatest percentage changes
                Dim percent_change As Double
                    percent_change = ((close_stock - open_stock) / open_stock)
                ' Condition check for greatest percentage increase
                If percent_change > greatest_incr Then
                    greatest_incr = percent_change
                    ticker_greatest_incr = ticker_sym
                ' Same but for greatest percentage decrease
                ElseIf percent_change < greatest_decr Then
                    greatest_decr = percent_change
                    ticker_greatest_decr = ticker_sym
                End If
                
                ' Check for greatest total volume
                If stock_vol > greatest_vol Then
                    greatest_vol = stock_vol
                    greatest_vol_ticker = ticker_sym
                End If
                
                ' Add one to the yearly output row
                yearly_output = yearly_output + 1
                
                ' Reset the total stock volume and opening stock value
                stock_vol = 0
                open_stock = ws.Cells(i + 1, 3).Value
            Else
                ' Add to the total stock volume
                stock_vol = stock_vol + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Format the output column for yearly change percentage to appear as percents
        Dim percent_format As Range
        Set percent_format = ws.Range("K:K")
        percent_format.NumberFormat = "0.00%"
        ' Format the colors to appear as green in yearly change unless they are negative
        Dim j As Long
        Dim last_j_output As Long
        With ActiveSheet
            last_j_output = ws.Cells(.Rows.Count, "J").End(xlUp).Row
        End With
        For j = 2 To last_j_output
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        ' Same formatting but for percent change
        Dim k As Long
        Dim last_k_output As Long
        With ActiveSheet
            last_k_output = ws.Cells(.Rows.Count, "K").End(xlUp).Row
        End With
        For k = 2 To last_k_output
            If ws.Cells(k, 11).Value >= 0 Then
                ws.Cells(k, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(k, 11).Interior.ColorIndex = 3
            End If
        Next k
        ' Output values for column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ' Output the results for greatest percentage increase
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ' Same for greatest percentage decrease
        ws.Range("P1").Value = "Ticker"
        ws.Range("P2").Value = ticker_greatest_incr
        ws.Range("P3").Value = ticker_greatest_decr
        ws.Range("P4").Value = greatest_vol_ticker
        ' Same for greatest total stock volume
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2").Value = greatest_incr
        ws.Range("Q3").Value = greatest_decr
        ws.Range("Q4").Value = greatest_vol
        ' Format greatest percentage increase and decrease to appear as percentages
        Set percent_format = ws.Range("Q2", "Q3")
        percent_format.NumberFormat = "0.00%"
        ' Auto fit formatting to all output columns
        ws.Columns("I:I").AutoFit
        ws.Columns("J:J").AutoFit
        ws.Columns("K:K").AutoFit
        ws.Columns("L:L").AutoFit
        ws.Columns("M:M").AutoFit
        ws.Columns("O:O").AutoFit
        ws.Columns("P:P").AutoFit
        ws.Columns("Q:Q").AutoFit
    Next ws
End Sub




