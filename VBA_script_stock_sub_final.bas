Attribute VB_Name = "Module1"
Option Explicit

Sub stock_sub_final()

Dim ws As Worksheet

    For Each ws In Worksheets

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        Dim i, NextRow, OpenIdx As Integer
        Dim YearlyChange, PercentChange As Double
        Dim StockTotal As LongLong
        Dim LastRow, TickerName As String
        Dim RowsK, RowsL As Integer
        Dim IdxB, IdxW, IdxM As Integer
        
        RowsK = 2
        RowsL = 2
        NextRow = 2
        OpenIdx = 2
        StockTotal = 0
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerName = ws.Cells(i, 1).Value
                ws.Range("I" & NextRow).Value = TickerName
                
                YearlyChange = ws.Cells(i, 6).Value - ws.Cells(OpenIdx, 3).Value
                ws.Range("J" & NextRow).Value = YearlyChange
                
                PercentChange = YearlyChange / ws.Cells(OpenIdx, 3).Value
                ws.Range("K" & NextRow).Value = PercentChange
                
                StockTotal = StockTotal + ws.Cells(i, 7).Value
                ws.Range("L" & NextRow).Value = StockTotal
                
                YearlyChange = 0
                PercentChange = 0
                StockTotal = 0
                OpenIdx = OpenIdx + 1
                NextRow = NextRow + 1
                
            Else
                StockTotal = StockTotal + ws.Cells(i, 7).Value
                
            End If
            
            '---------------------COLORINDEX---------------------
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 0
            End If
            
           '--------------------NUMBERFORMAT-------------------
            ws.Cells(i, 11).NumberFormat = "0.00%"
            ws.Cells(i, 10).NumberFormat = "0.00"
        
        Next i
       
            '---------------------------------------------------
            
        RowsK = ws.Cells(Rows.Count, "K").End(xlUp).Row
        RowsL = ws.Cells(Rows.Count, "L").End(xlUp).Row

        ws.Range("P2").Value = WorksheetFunction.Max(ws.Range("K2:K" & RowsK).Value)
        ws.Range("P3").Value = WorksheetFunction.Min(ws.Range("K2:K" & RowsK).Value)
        ws.Range("P4").Value = WorksheetFunction.Max(ws.Range("L2:L" & RowsL).Value)

        IdxB = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowsK).Value), ws.Range("K2:K" & RowsK).Value, 0)
        IdxW = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowsK).Value), ws.Range("K2:K" & RowsK).Value, 0)
        IdxM = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowsL).Value), ws.Range("L2:L" & RowsL).Value, 0)


        ws.Range("O2").Value = ws.Cells(IdxB + 1, 9).Value
        ws.Range("O3").Value = ws.Cells(IdxW + 1, 9).Value
        ws.Range("O4").Value = ws.Cells(IdxM + 1, 9).Value
    
        ws.Range("P2:P3").NumberFormat = "0.00%"
    
    
    Next ws


End Sub

