Sub StockAnalysis()


    'declaring variables

    Dim ws As Worksheet
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim TotalStockVolume As Double
    Dim PercentChange As Double
    Dim LastRow As Long
    Dim WorksheetName As String
    Dim Summary_Table_Row As Integer
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim i As Long


    For Each ws In Worksheets

        'running the macro across all the worksheets
        WorksheetName = ws.Name

        Summary_Table_Row = 2

        ws.Cells(1, 3).Value = "Yearly Open"

        ws.Cells(1, 6).Value = "Yearly Close"

        ws.Cells(1, 9).Value = "Ticker"

        ws.Cells(1, 10).Value = "Yearly Change"


        ws.Cells(1, 11).Value = " Percent Change"

        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Finding the last row of each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        MsgBox(LastRow)

        For i = 2 To LastRow

            ' check if stock changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' set currrent stock 
                ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(i , 1).Value
                Summary_Table_Row = Summary_Table_Row + 1
                'ws.Range("I" + Str(Summary_Table_Row) =  ws.Cells(i + 1, 1).Value
      
              
                TotalStockVolume = TotalStockVolume + ws.Cells(i , 7).Value
                ' set total value for the last total value for that stock
                TotalStockVolume = 0
             
            
                'set total value for new stock to zero
                 ws.Cells(Summary_Table_Row, 12).Value  = TotalStockVolume
                

                'YearlyOpen = ws.Cells(i + 1, 3).Value
                'YearlyClose = ws.Cells(i + 1, 6).Value
                'YearlyChange = (YearlyOpen - YearlyClose)
               
                'PercentChange = ((YearlyOpen - YearlyClose) / (YearlyOpen))
              
            Else
                 TotalStockVolume = TotalStockVolume + ws.Cells(i + 1, 7).Value
                ' YearlyChange = (YearlyOpen - YearlyClose)
                ' PercentChange = (YearlyChange / (YearlyOpen))
                
                
            End If

        Next i


        ' If ws.Cells(i + 1, 10).Value > 0 Then
        '     ws.Cells(1, 10).Interior.Color.Index = 4
            
        ' ElseIf ws.Cells(1 + 1, 10).Value > 0 Then
        '     ws.Range(j).InteriorColor.Index = 3
            
        ' End If

    Next ws

End Sub
