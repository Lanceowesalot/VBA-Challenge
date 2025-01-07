
Sub StockTicker()

    Dim ws As Worksheet
    Dim quarter_row As Integer
    Dim ticker As String
    Dim begin_price As Double
    Dim end_price As Double
    Dim change_amount As Double
    Dim percent_change As Double
    Dim stock_volume As Double
    Dim ticker_percent_increase As String
    Dim ticker_percent_decrease As String
    Dim ticker_volume As String
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease As Double
    Dim greatest_volume As Double
    Dim i As Long
    Dim lastRowA As Long
    

    ' Worksheet Loop for all tabs
    For Each ws In Worksheets
        On Error Resume Next
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If lastRowA < 2 Then
           
        End If
      
        ' Add column and row headers
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        ws.Range("P1:Q1").Value = Array("Ticker", "Value")
        ws.Range("O2:O4").Value = ws.Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))

        ' Initialize variables
        quarter_row = 2
        ticker = ws.Cells(2, "A").Value
        begin_price = ws.Cells(2, "C").Value
        stock_volume = 0
        greatest_percent_increase = 0
        greatest_percent_decrease = 0
        greatest_volume = 0

        ' Processing rows
        For i = 2 To lastRowA
            ' add to stock volume
            stock_volume = stock_volume + ws.Cells(i, "G").Value
            
            ' Check for change in ticker or last row
            If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Or i = lastRowA Then
                ' Calculate percent and yearly change
                end_price = ws.Cells(i, "F").Value
                change_amount = end_price - begin_price
                If begin_price <> 0 Then
                    percent_change = change_amount / begin_price
                Else
                    percent_change = 0
                End If

                ' collate data to analysis columns
                ws.Cells(quarter_row, "I").Value = ticker
                ws.Cells(quarter_row, "J").Value = change_amount
                If change_amount < 0 Then
                    ws.Cells(quarter_row, "J").Interior.ColorIndex = 3 ' Red for negative
                End If
                If change_amount > 0 Then
                    ws.Cells(quarter_row, "J").Interior.ColorIndex = 4 ' Green for positive
                End If
                If change_amount = 0 Then
                    ws.Cells(quarter_row, "J").Interior.ColorIndex = 2 ' White for no change
                End If
                
                ws.Cells(quarter_row, "K").Value = percent_change
                ws.Cells(quarter_row, "K").NumberFormat = "0.00%"
                ws.Cells(quarter_row, "L").Value = stock_volume

                ' Update greatest values
                If quarter_row = 2 Or percent_change > greatest_percent_increase Then
                    greatest_percent_increase = percent_change
                    ticker_percent_increase = ticker
                End If
                If quarter_row = 2 Or percent_change < greatest_percent_decrease Then
                    greatest_percent_decrease = percent_change
                    ticker_percent_decrease = ticker
                End If
                If quarter_row = 2 Or stock_volume > greatest_volume Then
                    greatest_volume = stock_volume
                    ticker_volume = ticker
                End If

                ' Ticker reset
                quarter_row = quarter_row + 1
                If i < lastRowA Then
                    ticker = ws.Cells(i + 1, "A").Value
                    begin_price = ws.Cells(i + 1, "C").Value
                    stock_volume = 0
                End If
            End If
        Next i

        ' Output greatest values
        ws.Cells(2, "P").Value = ticker_percent_increase
        ws.Cells(2, "Q").Value = greatest_percent_increase
        ws.Cells(2, "Q").NumberFormat = "0.00%"
        ws.Cells(3, "P").Value = ticker_percent_decrease
        ws.Cells(3, "Q").Value = greatest_percent_decrease
        ws.Cells(3, "Q").NumberFormat = "0.00%"
        ws.Cells(4, "P").Value = ticker_volume
        ws.Cells(4, "Q").Value = greatest_volume

        ' Adjust column widths
        ws.Columns("A:Q").AutoFit

    Next ws

End Sub
